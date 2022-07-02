VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "贪吃蛇"
   ClientHeight    =   12165
   ClientLeft      =   5280
   ClientTop       =   1650
   ClientWidth     =   18975
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleMode       =   0  'User
   ScaleWidth      =   18975
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   101
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   100
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   99
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   98
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   97
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   96
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   95
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   94
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   93
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   92
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   91
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   90
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   89
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   88
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   87
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   86
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   85
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   84
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   83
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   82
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   81
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   80
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   79
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   78
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   77
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   76
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   75
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   74
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   73
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   72
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   71
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   70
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   69
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   68
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   67
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   66
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   65
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   64
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   63
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   62
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   61
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   60
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   59
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   58
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   57
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   56
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   55
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   54
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   53
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   52
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   51
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   50
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   49
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   48
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   47
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   46
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   45
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   44
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   43
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   42
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   41
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   40
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   39
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   38
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   37
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   36
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   35
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   34
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   33
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   32
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   31
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   28
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   120
      Top             =   2520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "请开始你的表演"
      Height          =   735
      Left            =   14280
      TabIndex        =   2
      Top             =   11400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   400
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   255
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

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
Command3.SetFocus
End Sub

Private Sub Command3_Click()
Timer1 = Not Timer1
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

Private Sub Form_Load()
Rem maxleng,initialleng为标准模块中声明的全局变量
For i = i To maxleng
Load Command1(Command1.UBound + 1)
Next i

'form1.width=19218 height=12750,command1.width height=255,command2.height width=400
Form1.BorderStyle = 1
Command3.Caption = "开始"
For i = 1 To Command1.UBound
 Command1(i).BackColor = RGB(i / Command1.UBound * 200 + 55, i / Command1.UBound * 200 + 55, i / Command1.UBound * 200 + 55)
 Command1(i).Width = Command1(0).Width
 Command1(i).Height = Command1(0).Height
 Command1(i).Caption = ""
 If i > initialleng Then
 Command1(i).Visible = False
 Else
 Command1(i).Visible = True
 Command1(i).Top = Command1(0).Top
 Command1(i).Left = Command1(0).Left - i * Command1(0).Width
 End If
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

End Sub

Private Sub Timer1_Timer()
Randomize

For i = 1 To Command1.UBound
  If Command1(i).Visible = False Then Exit For
Next i

st = i - 1
If st = Command1.UBound Then MsgBox "你是针不戳", , "拜拜了你嘞": End   '成功

Rem 移动
For i = st To 1 Step -1
Command1(i).Left = Command1(i - 1).Left
Command1(i).Top = Command1(i - 1).Top
Next i
If w Then Command1(0).Top = Command1(0).Top - Command1(0).Height
If s Then Command1(0).Top = Command1(0).Top + Command1(0).Height
If a Then Command1(0).Left = Command1(0).Left - Command1(0).Width
If d Then Command1(0).Left = Command1(0).Left + Command1(0).Width
Rem 死忙
If Command1(0).Top <= 0 Or Command1(0).Top + Command1(0).Height >= Form1.ScaleHeight Or Command1(0).Left <= 0 Or Command1(0).Left + Command1(0).Width >= Form1.ScaleWidth Then MsgBox "你碰壁自杀了", , "一首凉凉送给你": End
For i = 1 To Command1.UBound
 If Command1(0).Top = Command1(i).Top And Command1(0).Left = Command1(i).Left Then MsgBox "你把自己咬死了", , "一首凉凉送给你": End
Next i
Rem 成长
If ((Command1(0).Left >= Command2.Left And Command1(0).Left <= Command2.Left + Command2.Width) Or (Command1(0).Left + Command1(0).Width >= Command2.Left And Command1(0).Left + Command1(0).Width <= Command2.Left + Command2.Width)) And ((Command1(0).Top >= Command2.Top And Command1(0).Top <= Command2.Top + Command2.Height) Or (Command1(0).Top + Command1(0).Height >= Command2.Top And Command1(0).Top + Command1(0).Height <= Command2.Top + Command2.Height)) Then
Do
  Command2.Top = Rnd * Form1.ScaleHeight
  Command2.Left = Rnd * Form1.ScaleWidth
 Loop Until Command2.Top + Command2.Height < Form1.ScaleHeight And Command2.Left + Command2.Width < Form1.ScaleWidth
 Command2.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
 Command1(st + 1).Visible = True
 Command1(st + 1).Left = Command1(st).Left
 Command1(st + 1).Top = Command1(st).Top
End If
End Sub

