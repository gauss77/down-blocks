VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJogo 
   Caption         =   "Down Blocks"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9720
   Icon            =   "frmJogo.frx":0000
   LinkTopic       =   "frmJogo"
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerMovimento 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   4320
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Depuração do Sistema "
      Height          =   6855
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtDebug 
         Appearance      =   0  'Flat
         Height          =   3615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton cmdDebugEsconder 
         Caption         =   "Finalizar Depuração"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label lblDebugMatrizTexto 
         Caption         =   "Matriz ""Jogo"""
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pctJogo 
      BackColor       =   &H00C0C0C0&
      Height          =   6810
      Left            =   120
      ScaleHeight     =   6750
      ScaleWidth      =   3750
      TabIndex        =   7
      Top             =   120
      Width           =   3810
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   179
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   178
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   177
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   176
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   175
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   174
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   173
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   172
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   171
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   170
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   169
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   168
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   167
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   166
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   165
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   164
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   163
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   162
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   161
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   160
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   159
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   158
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   157
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   156
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   155
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   154
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   153
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   152
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   151
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   150
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   149
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   148
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   147
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   146
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   145
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   144
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   143
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   142
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   141
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   140
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   139
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   138
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   137
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   136
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   135
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   134
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   133
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   132
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   131
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   130
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   129
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   128
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   127
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   126
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   125
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   124
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   123
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   122
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   121
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   120
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   119
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   118
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   117
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   116
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   115
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   114
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   113
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   112
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   111
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   110
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   109
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   108
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   107
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   106
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   105
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   104
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   103
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   102
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   101
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   100
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   99
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   98
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   97
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   96
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   95
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   94
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   93
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   92
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   91
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   90
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   89
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   88
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   87
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   86
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   85
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   84
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   83
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   82
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   81
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   80
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   79
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   78
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   77
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   76
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   75
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   74
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   73
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   72
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   71
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   70
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   69
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   68
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   67
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   66
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   65
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   64
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   63
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   62
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   61
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   60
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   59
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   58
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   57
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   56
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   55
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   54
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   53
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   52
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   51
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   50
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   49
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   48
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   47
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   46
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   45
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   44
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   43
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   42
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   41
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   40
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   39
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   38
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   37
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   36
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   35
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   34
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   33
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   32
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   31
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   30
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   29
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   28
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   27
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   26
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   25
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   24
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   23
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   22
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   21
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   20
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   19
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   18
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   17
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   16
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   12
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   8
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBloco 
         Height          =   375
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImageListBlocos 
      Left            =   5400
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer TimerJogo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   4320
   End
   Begin VB.CommandButton cmdNovoJogo 
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   6000
      Width           =   1455
   End
   Begin VB.PictureBox pctProx 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   4070
      ScaleHeight     =   870
      ScaleWidth      =   1830
      TabIndex        =   1
      Top             =   600
      Width           =   1860
      Begin VB.Image imgBlocoProx 
         Height          =   375
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBlocoProx 
         Height          =   375
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBlocoProx 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgBlocoProx 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Line Line4 
      X1              =   4080
      X2              =   5880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblTXTNivel 
      Alignment       =   2  'Center
      Caption         =   "Nível"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblPontos 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   5880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblTXTPontos 
      Alignment       =   2  'Center
      Caption         =   "Pontuação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   5880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   5880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblTXTProx 
      Alignment       =   2  'Center
      Caption         =   "Próximo Bloco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu mnuJogo 
      Caption         =   "Jogo"
      Begin VB.Menu mnuNovoJogo 
         Caption         =   "Novo Jogo"
      End
      Begin VB.Menu mnuFinalizar 
         Caption         =   "Finalizar Jogo"
      End
      Begin VB.Menu mnuEspaco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuOpcoes 
      Caption         =   "Opções"
      Begin VB.Menu mnuMusica 
         Caption         =   "Música"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSons 
         Caption         =   "Sons"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEspaco3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEstiloTexto 
         Caption         =   "Estilo do Bloco"
         Begin VB.Menu mnuEst 
            Caption         =   "Clássico"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuEst 
            Caption         =   "Novo"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEspaco4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNivelTexto 
         Caption         =   "Nível"
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 0"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 1"
            Index           =   1
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 2"
            Index           =   2
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 3"
            Index           =   3
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 4"
            Index           =   4
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 5"
            Index           =   5
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 6"
            Index           =   6
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 7"
            Index           =   7
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 8"
            Index           =   8
         End
         Begin VB.Menu mnuNivel 
            Caption         =   "Nível 9"
            Index           =   9
         End
      End
      Begin VB.Menu mnuEspaco5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIdiomaTexto 
         Caption         =   "Idioma"
         Begin VB.Menu mnuIdioma 
            Caption         =   "Português"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuIdioma 
            Caption         =   "English"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Ajuda"
      Begin VB.Menu mnuConteudo 
         Caption         =   "Conteúdo"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug"
      End
      Begin VB.Menu mnuEspaco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "Sobre"
      End
   End
End
Attribute VB_Name = "frmJogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TeclaPressionada As Integer 'Indica a tecla que foi pressionada

Dim cmdNovoJogoSTATUS As Integer 'Indica qual a posição da va-
                                 'riável "cmdNovoJogoTEXTO"
                                 'está sendo usada no momento
                                 'como rótulo em "cmdNovoJogo"

Dim TipoBlocoProx As Integer, EstiloBlocoProx As Integer ' Indicam o Tipo
                                                         'e o Estilo dos
                                                         'Blocos gerados
                                                         'em "pctProx"

Dim TipoBlocoEmJogo As Integer ' Armazena o Tipo do Bloco
                               'que estiver em primeiro plano
                               'no jogo (Bloco que está apto
                               'a ser movido em "pctJogo")
                               
Dim PosicaoDoBloco As Integer ' Indica as possíveis posições de movimentação
                              'de cada Bloco, a saber:
                   ' Obs.: Os blocos serão representados pelos
                   'números que se referenciam a cada posição de
                   'cada bloco no conjunto
                   '
                   '=============================
                   '=TIPO BLOCO: 0 ( **** )     =
                   '=EIXO: 2                    =
                   '=  PosicaoDoBloco = 0: 1234 =
                   '=                           =
                   '=  PosicaoDoBloco = 1: 1    =
                   '=                      2    =
                   '=                      3    =
                   '=                      4    =
                   '=============================
                   '
                   '=============================
                   '=TIPO BLOCO: 1 ( *** )      =
                   '=EIXO: 2                    =
                   '=  PosicaoDoBloco = 0: 123  =
                   '=                           =
                   '=  PosicaoDoBloco = 1: 1    =
                   '=                      2    =
                   '=                      3    =
                   '=============================
                   '
                   '=============================
                   '=                 *         =
                   '=TIPO BLOCO: 2 ( *** )      =
                   '=EIXO: 3                    =
                   '=                       1   =
                   '=  PosicaoDoBloco = 0: 234  =
                   '=                           =
                   '=  PosicaoDoBloco = 1:  4   =
                   '=                      13   =
                   '=                       2   =
                   '=                           =
                   '=  PosicaoDoBloco = 2: 432  =
                   '=                       1   =
                   '=                           =
                   '=  PosicaoDoBloco = 3: 2    =
                   '=                      31   =
                   '=                      4    =
                   '=============================
                   '
                   '=============================
                   '=                  *        =
                   '=TIPO BLOCO: 3 ( *** )      =
                   '=EIXO: 3                    =
                   '=                        1  =
                   '=  PosicaoDoBloco = 0: 234  =
                   '=                           =
                   '=  PosicaoDoBloco = 1: 14   =
                   '=                       3   =
                   '=                       2   =
                   '=                           =
                   '=  PosicaoDoBloco = 2: 432  =
                   '=                      1    =
                   '=                           =
                   '=  PosicaoDoBloco = 3: 2    =
                   '=                      3    =
                   '=                      41   =
                   '=============================
                   '
                   '=============================
                   '=                 **        =
                   '=TIPO BLOCO: 4 ( **  )      =
                   '=EIXO: 4                    =
                   '=                       12  =
                   '=  PosicaoDoBloco = 0: 34   =
                   '=                           =
                   '=  PosicaoDoBloco = 1: 1    =
                   '=                      24   =
                   '=                       3   =
                   '=============================
                   '
                   '=============================
                   '=                **         =
                   '=TIPO BLOCO: 5 ( ** )       =
                   '=EIXO: < Não tem >          =
                   '=                      12   =
                   '=  PosicaoDoBloco = 0: 34   =
                   '=============================
                   
    'VARIÁVEIS de uso nos eventos de "TimerJogo"
    Dim ContadorTimerJogo As Integer
    Dim Blocos_a_VerificarTEMP As PosicoesParaIndices
    Dim Blocos_a_Verificar(4) As Integer
    Dim Linha_Completa As Integer
    Dim CriarProxBloco As Boolean
    Dim Linha_da_colisao As Integer
    
    'VARIÁVEIS de uso nos eventos de "TimerMovimento"
    Dim Blocos_a_AnalisarTEMP As PosicoesParaIndices
    Dim Blocos_a_Analisar(4) As Integer
    Dim ContadorTimerMovimento As Integer
    Dim ColidiuEsquerdaDireita As Boolean

Private Sub cmdDebugEsconder_Click()
'Esconde a parte do formulário com os controles de depuração

    frmJogo.Width = 6120
    frmJogo.Left = frmJogo.Left + 1860

End Sub

Private Sub cmdNovoJogo_Click()
    
    'Inicia o Jogo
    IniciarJogo

End Sub

Private Sub Command1_Click()
GerarBlocos "pctJogo", 0, 2
GerarBlocos "pctProx", 4, 5



MoverBloco 37, IndicesEmJogo(1), IndicesEmJogo(2), IndicesEmJogo(3), IndicesEmJogo(4)

Dim a As Posicao

a = LocalizarBlocoNaMatriz(2)

MsgBox "X: " & a.PosicaoX & " Y:" & a.PosicaoY

End Sub

Private Sub pctJogo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    TeclaPressionada = KeyCode

End Sub

Private Sub TimerMovimento_Timer()
'Realiza os movimentos laterais dos Blocos
    
    ' Verifica se deve haver movimento para esquerda, direita,
    'mudança de posição do Bloco ou aceleração do movimento
    Select Case TeclaPressionada
    
        Case 38 'Seta para cima
            'Muda as posições do Bloco
            
        
        Case 40 'Seta para baixo
            'Acelera a descida do Bloco
            
            'TimerJogo.Interval = 1

   
        Case 37 'Seta para o lado esquerdo
            'Move o Bloco uma posição para a esquerda
            
            Blocos_a_AnalisarTEMP = ArmazenarPosicoes(37, TipoBlocoEmJogo, PosicaoDoBloco)
            Blocos_a_Analisar(0) = Blocos_a_AnalisarTEMP.idx1
            Blocos_a_Analisar(1) = Blocos_a_AnalisarTEMP.idx2
            Blocos_a_Analisar(2) = Blocos_a_AnalisarTEMP.idx3
            Blocos_a_Analisar(3) = Blocos_a_AnalisarTEMP.idx4

            GoTo MoverEsquerdaDireita

        Case 39 'Seta para o lado direito
            'Move o Bloco uma posição para a direita
            
            Blocos_a_AnalisarTEMP = ArmazenarPosicoes(39, TipoBlocoEmJogo, PosicaoDoBloco)
            Blocos_a_Analisar(0) = Blocos_a_AnalisarTEMP.idx1
            Blocos_a_Analisar(1) = Blocos_a_AnalisarTEMP.idx2
            Blocos_a_Analisar(2) = Blocos_a_AnalisarTEMP.idx3
            Blocos_a_Analisar(3) = Blocos_a_AnalisarTEMP.idx4

            GoTo MoverEsquerdaDireita

    End Select
    
    TeclaPressionada = 0

    Exit Sub
    
MoverEsquerdaDireita:
    
    'Analisa as posições solicitadas do Bloco em Jogo
    ' Obs.: A informação "999" armazenada em "Blocos_a_Analisar"
    'indica que aquela posição não deve ser verificada
    ContadorTimerMovimento = 0
    
    ColidiuEsquerdaDireita = False
    
    Do While ContadorTimerMovimento <= 3
    
        If Blocos_a_Analisar(ContadorTimerMovimento) <> 999 Then
    
            'Verifica:
            '   Se movimento é para ESQUERDA:
            '       - Se é coluna 1; sendo, não move o bloco
            '       - Se há algum bloco imediatamente a ESQUERDA
            '        deste; havendo, não move o bloco
            '   Se movimento é para DIREITA:
            '       - Se é coluna 18; sendo, não move o bloco
            '       - Se há algum bloco imediatamente a DIREITA
            '        deste; havendo, não move o bloco
            
            Select Case TeclaPressionada
        
                Case 37 'Lado esquerdo
                    'Verifica se há colisão
                    If DetectarColisao(IndicesEmJogo(Blocos_a_Analisar(ContadorTimerMovimento)), 37) = True Then

                         ColidiuEsquerdaDireita = True
                
                        Exit Do
                
                    End If
                
                Case 39 'Lado direito
                    'Verifica se há colisão
                    If DetectarColisao(IndicesEmJogo(Blocos_a_Analisar(ContadorTimerMovimento)), 39) = True Then

                        ColidiuEsquerdaDireita = True
                        
                        Exit Do
                
                    End If
                
            End Select

        
        End If
    
        ContadorTimerMovimento = ContadorTimerMovimento + 1
    
    Loop
    
    'Move os blocos
    If ColidiuEsquerdaDireita = False Then MoverBloco _
      TeclaPressionada, IndicesEmJogo(1), IndicesEmJogo(2), _
      IndicesEmJogo(3), IndicesEmJogo(4)
    
    TeclaPressionada = 0

End Sub

Private Sub TimerJogo_Timer()
'Realiza o movimento de descida dos Blocos
    
    CriarProxBloco = False
    
    Blocos_a_VerificarTEMP = ArmazenarPosicoes(40, TipoBlocoEmJogo, PosicaoDoBloco)
    Blocos_a_Verificar(0) = Blocos_a_VerificarTEMP.idx1
    Blocos_a_Verificar(1) = Blocos_a_VerificarTEMP.idx2
    Blocos_a_Verificar(2) = Blocos_a_VerificarTEMP.idx3
    Blocos_a_Verificar(3) = Blocos_a_VerificarTEMP.idx4
    
    'Verifica as posições de blocos solicitadas
    ' Obs.: A informação "999" armazenada em "Blocos_a_Verificar"
    'indica que aquela posição não deve ser verificada
    ContadorTimerJogo = 0
    
    Do While ContadorTimerJogo <= 3
    
        If Blocos_a_Verificar(ContadorTimerJogo) <> 999 Then
        
            'Verifica se há colisão
            If DetectarColisao(IndicesEmJogo(Blocos_a_Verificar(ContadorTimerJogo)), 40) = True Then

                CriarProxBloco = True
                
                Exit Do
                
            End If
        
        End If
    
        ContadorTimerJogo = ContadorTimerJogo + 1
    
    Loop
    
    ' Se "CriarProxBloco = True", houve colisão. Assim,
    'primeiramente verifica-se se há alguma linha completa
    'na matriz e, caso haja, exclui-se esta linha e desce-se
    'todos os blocos da matriz até a posição ocupada anteriormente
    'por esta linha. Depois, gera-se um novo Bloco.
    If CriarProxBloco = True Then

        'Toca o som de colisão do bloco
        If mnuSons.Checked = True Then
                    
            TocarSom App.Path & "\Sons\Parada.wav", SND_ASYNC
                            
        End If

VerificarNovamente:

        'Verifica se alguma linha está completa
        Linha_Completa = VerificarLinhasCompletas

        If Linha_Completa <> 999 Then
        
            'Se alguma linha estiver completa:
            '   - marca-se os blocos desta linha como não mais em
            '    uso;
            '   - esconde-se os blocos desta linha;
            '   - movem-se todos os blocos da linha superior a
            '    esta para esta posição.
            
            'Aumenta-se em 100 a pontuação
            lblPontos.Caption = CStr(CDbl(lblPontos.Caption) + 100)
            
            ContadorTimerJogo = 1
            
            Do While ContadorTimerJogo <= 10
            
                'Marca o bloco como não mais em uso
                BlocoEmJogo(Jogo(Linha_Completa, ContadorTimerJogo)) = False
                
                'Esconde-se o bloco, retornando-o à posição inicial
                With imgBloco(Jogo(Linha_Completa, ContadorTimerJogo))
                    .Top = 0
                    .Left = 0
                    .Visible = False
                    .Picture = Nothing
                End With
                
                ' Armazena "999" em "EstiloBlocoEmJogo", indicando que não há
                'figura neste bloco (o bloco não está em uso)
                EstiloBlocoEmJogo(Jogo(Linha_Completa, ContadorTimerJogo)) = 999
                
                'Armazena 999 na posição deste bloco na matriz "Jogo"
                Jogo(Linha_Completa, ContadorTimerJogo) = 999
            
                ContadorTimerJogo = ContadorTimerJogo + 1
            
            Loop

            'Reordena os blocos em "pctJogo" e os dados da matriz
            ReordenarJogo (Linha_Completa)
            
            'Verifica novamente se não há mais linhas completas
            GoTo VerificarNovamente
        
        End If
           
        'Gera os Blocos que estavam em "pctProx" em "pctJogo"...
        GerarBlocos "pctJogo", TipoBlocoProx, EstiloBlocoProx
        
        TipoBlocoEmJogo = TipoBlocoProx
        
        '...e depois gera novos Blocos em "pctProx"
        TipoBlocoProx = Random(5, (5 * Second(Time)))
        EstiloBlocoProx = Random(7, (7 * Second(Time)))
        
        GerarBlocos "pctProx", TipoBlocoProx, EstiloBlocoProx
    
        'Indica que a posição atual do Bloco é a inicial ("0")
        PosicaoDoBloco = 0
        
        'AFERIÇÃO DO FIM DE JOGO
        
        'Verifica se as seguintes posições na matriz estão ocupadas:
        'PARA OS BLOCOS TIPOS "1" e "2":
        '   - (2,4);
        '   - (2,5);
        '   - (2,6).
        'PARA OS DEMAIS BLOCOS:
        '   - (3,4);
        '   - (3,5);
        '   - (3,6).
        'Havendo blocos nesta posição, indica-se fim-de-jogo
        
        ' Verifica o tipo do novo Bloco, que definirá qual linha
        'da matriz deverá ser verificada para descobrir se há
        'blocos em certas posições
        If TipoBlocoEmJogo = 0 Or TipoBlocoEmJogo = 1 Then
      
            Linha_da_colisao = 2
                
         Else
            
            Linha_da_colisao = 3
                
        End If
        
        ContadorTimerJogo = 4
        
        Do While ContadorTimerJogo <= 6
'FIM DE JOGO ===================================
            If Jogo(Linha_da_colisao, ContadorTimerJogo) <> 999 Then
                'Fim de jogo
                MsgBox "Desculpe, fim de jogo..."
                TimerJogo.Enabled = False
                TimerMovimento.Enabled = False
                
                'O menu de término de jogo é desabilitado
                mnuFinalizar.Enabled = False
                
                'Atribue os valores iniciais aos rótulos
                lblPontos.Caption = "0"
                cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
                
                'Habilita os menus selecionadores de Nível de Jogo
                MenuNivel ("Habilitar")
                
                Exit Sub
                
            End If
        
            ContadorTimerJogo = ContadorTimerJogo + 1
        
        Loop
    
    Else
    
        'Move os blocos uma posição para baixo
        MoverBloco 40, IndicesEmJogo(1), IndicesEmJogo(2), IndicesEmJogo(3), IndicesEmJogo(4)
    
        'Toca o som de descida do bloco
        If mnuSons.Checked = True Then
    
            TocarSom App.Path & "\Sons\Descendo.wav", SND_ASYNC
            
        End If
        
    End If
    
'USO NA DEPURAÇÃO DO SISTEMA ********************************
    
    'Atualiza a exibição da matriz "Jogo" em "txtDebug"
    ExibirValJogo
    
'************************************************************

End Sub

Private Sub Form_Load()
' Logo ao iniciar, carrega o ImageList com o Estilo de Bloco
'"Clássico", e define o nível inicial como 1 (1000 milissegundos)

    CarregarImageList (App.Path & "\Blocos\Clássico\")
    
    'Indica o nível do Jogo (Inicialmente "0" - 1000 milissegundos)
    SelecionarNivel 0, False
    
    'Carrega os rótulos de "cmdNovoJogo" na variável
    '"cmdNovoJogoTEXTO"
    cmdNovoJogoTEXTO(0) = "Iniciar"
    cmdNovoJogoTEXTO(1) = "Pausar"
    cmdNovoJogoTEXTO(2) = "Continuar"
    
    'Indica o rótulo utilizado em "cmdNovoJogo" (no caso a
    'posição "0" da variável "cmdNovoJogoTEXTO"
    cmdNovoJogoSTATUS = 0
    
    ' Armazena a constante que define uma posição na movimentação
    'dos blocos em "pctJogo"
    UmaPosicao = 375
    
    'O menu de término de jogo é desabilitado
    mnuFinalizar.Enabled = False
    
    'Prepara o jogo para ser iniciado
    PrepararJogo
    
'COMANDOS PARA DEPURAÇÃO DO SISTEMA ************************
            
    'Esconde a parte do formulário com os controles de depuração
    frmJogo.Width = 6120
    
    'Já cria uma primeira imagem da variável "Jogo"
    ExibirValJogo
    
'***********************************************************

End Sub

'===========================================================
'COMANDOS REFERENTES AOS MENUS                             =
'===========================================================
Private Sub mnuNovoJogo_Click()
'Limpa a tela de Jogo e, depois, inicia um
'novo Jogo, habilitando o Timer ("TimerJogo")
    
    'Atribue os valores iniciais aos rótulos
    lblPontos.Caption = "0"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
    
    'O menu de término de jogo é habilitado
    mnuFinalizar.Enabled = True
    
    'Prepara o jogo para ser iniciado
    PrepararJogo
    
    'Inicia o Jogo
    IniciarJogo
    
End Sub

Private Sub mnuFinalizar_Click()
'Finaliza o Jogo que estiver em andamento

    TimerJogo.Enabled = False
    TimerMovimento.Enabled = False
                
    'O menu de término de jogo é desabilitado
    mnuFinalizar.Enabled = False
                
    'Atribue os valores iniciais aos rótulos
    lblPontos.Caption = "0"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
                
    'Habilita os menus selecionadores de Nível de Jogo
    MenuNivel ("Habilitar")
    
    'Prepara o jogo para ser novamente inicializado
    PrepararJogo

    'O menu de término de jogo é desabilitado
    mnuFinalizar.Enabled = False

End Sub

Private Sub mnuSair_Click()
'Sai do jogo
    
    End

End Sub

Private Sub mnuDebug_Click()
'Exibe os controles para depuração do Jogo
    
    If frmJogo.Width = 6120 Then
    
        frmJogo.Width = 9840
        frmJogo.Left = frmJogo.Left - 1860
    
    Else
    
        frmJogo.Width = 6120
        frmJogo.Left = frmJogo.Left + 1860
    
    End If
    
End Sub

Private Sub mnuSobre_Click()
'Exibe "frmSobre"
    frmSobre.Show (vbModal)

End Sub

Private Sub mnuEst_Click(Index As Integer)
' Seleciona o Estilo do Bloco, conforme a seleção do usuário
'(indicada pela variável "Index")
    
    Dim ContadorEstilo As Integer
    
    'Tira o "Checked" dos menus de Estilos
    Uncheck ("Estilo")
    
    Select Case Index
    
        Case 0
        'Muda o Estilo dos blocos para "Clássico"
            
            mnuEst(0).Checked = True
            
            CarregarImageList (App.Path & "\Blocos\Clássico\")
        
        Case 1
        'Muda o Estilo dos blocos para "Clássico"
                 
            mnuEst(1).Checked = True
            
            CarregarImageList (App.Path & "\Blocos\Novo\")
        
    End Select
    
    ' Insere este novo estilo em todos os blocos que
    'já estiverem em jogo:
    
    '   - Em "pctJogo"
    ContadorEstilo = 0
    
    Do While ContadorEstilo < 180
    
        If EstiloBlocoEmJogo(ContadorEstilo) <> 999 Then
    
            imgBloco(ContadorEstilo).Picture = ImageListBlocos.ListImages(EstiloBlocoEmJogo(ContadorEstilo)).Picture
        
        End If
    
        ContadorEstilo = ContadorEstilo + 1
    
    Loop
    
    '   - Em "pctProx"
    ContadorEstilo = 0
    
    Do While ContadorEstilo < 4
    
        If imgBlocoProx(ContadorEstilo).Top <> 0 Then
    
            imgBlocoProx(ContadorEstilo).Picture = ImageListBlocos.ListImages(EstiloBlocoProx).Picture

        End If
    
        ContadorEstilo = ContadorEstilo + 1
    
    Loop

End Sub

Private Sub mnuNivel_Click(Index As Integer)
' Seleciona o Nível do Jogo (velocidade de descida dos Blocos),
'conforme a seleção do usuário (indicada pela variável "Index")
    
    'Tira o "Checked" dos menus de Níveis
    Uncheck ("Nível")
    
    SelecionarNivel Index, False

End Sub

Private Sub mnuIdioma_Click(Index As Integer)
'Troca o idioma do Jogo, utilizando como base a variável
'"Index", através da função "TrocarIdioma()"

    'Tira o "Checked" dos menus de Idioma
    Uncheck ("Idioma")

    Select Case Index
    
        Case 0
            mnuIdioma(0).Checked = True
            TrocarIdioma ("Ptb")
        
        Case 1
            mnuIdioma(1).Checked = True
            TrocarIdioma ("Eng")
            
    End Select
    
    'Coloca o novo Rótulo em "cmdNovoJogo"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(cmdNovoJogoSTATUS)
    
End Sub

Private Sub mnuMusica_Click()
' Esta SubRotina habilita ou não as músicas do Jogo
'Indica "True" no método Checked do menu selecionado (caso
'seja selecionada a opção) ou "False (caso a opção não seja
'selecionada)

    If mnuMusica.Checked = True Then

        mnuMusica.Checked = False
        
    Else
    
        mnuMusica.Checked = True
    
    End If

End Sub

Private Sub mnuSons_Click()
' Esta SubRotina habilita ou não os sons do Jogo
' Indica "True" no método Checked do menu selecionado (caso
'seja selecionada a opção) ou "False (caso a opção não seja
'selecionada)

    If mnuSons.Checked = True Then

        mnuSons.Checked = False
        
    Else
    
        mnuSons.Checked = True
    
    End If

End Sub
'===========================================================

Function LimpaTela()
'Limpa a tela de Jogo

    Dim ContadorNovoJogo1, ContadorNovoJogo2 As Integer
        
    'Armazena "999" em todas as posições da matriz "Jogo"
    ContadorNovoJogo1 = 1
    ContadorNovoJogo2 = 1

    Do While ContadorNovoJogo1 <= 18
    
        Do While ContadorNovoJogo2 <= 10
        
            Jogo(ContadorNovoJogo1, ContadorNovoJogo2) = 999

            ContadorNovoJogo2 = ContadorNovoJogo2 + 1
        
        Loop
        
        ContadorNovoJogo2 = 1
        
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop
    
'USO NA DEPURAÇÃO DO SISTEMA ********************************
    
    'Atualiza a exibição da matriz "Jogo" em "txtDebug"
    ExibirValJogo
    
'************************************************************
        
    'Limpa as variáveis
    TipoBlocoProx = 0
    EstiloBlocoProx = 0
    NumBloco = 0
    
    'Atribue os valores iniciais aos rótulos
    lblPontos.Caption = "0"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
    
    'Volta os blocos às posições iniciais e sem figuras anexas
    ContadorNovoJogo1 = 0
    
    Do While ContadorNovoJogo1 < 180
    
        With imgBloco(ContadorNovoJogo1)
            .Top = 0
            .Left = 0
            .Visible = False
            .Picture = Nothing
        End With
    
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop
    
    'Armazena "False" para todas as posições dos blocos de "pctJogo"
    ContadorNovoJogo1 = 0
    
    Do While ContadorNovoJogo1 < 180
    
        BlocoEmJogo(ContadorNovoJogo1) = False
    
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop

End Function

Function IniciarJogo()

'Inicia um novo Jogo, habilitando o Timer ("TimerJogo")
    
    If cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0) Then
        'Inicia o Jogo
        
        PrepararJogo
        
        'Limpa a tela do Jogo
        LimpaTela

        Dim TipoBlocoNovo As Integer, EstiloBlocoNovo As Integer
        
        ' Gera aleatoriamente as especificações dos Blocos de
        '"pctProx"
        TipoBlocoProx = Random(5, (5 * Second(Time)))
        EstiloBlocoProx = Random(7, (7 * Second(Time)))
        
        ' Usa como base de especificações do primeiro conjunto
        'de Blocos de "pctJogo" os valores de "TipoBlocoProx"
        'e "EstiloBlocoProx", acrescidos de 1 (Caso estes
        'valores extravasem "4" para "TipoBlocoNovo" e "7"
        'para "EstiloBlocoProx", um IF encarrega-se de manter
        'estes valores no intervalo necessário).
        TipoBlocoNovo = TipoBlocoProx + 1
        If TipoBlocoNovo = 6 Then TipoBlocoNovo = 4
        EstiloBlocoNovo = EstiloBlocoProx + 1
        If EstiloBlocoNovo = 8 Then EstiloBlocoNovo = 6

        'Armazena o Tipo do Bloco que está em jogo
        TipoBlocoEmJogo = TipoBlocoNovo

        'Gera os primeiros Blocos em "pctJogo"...
        GerarBlocos "pctJogo", TipoBlocoNovo, EstiloBlocoNovo

        '...e depois em "pctProx"
        
        GerarBlocos "pctProx", TipoBlocoProx, EstiloBlocoProx

        'Indica que a posição atual do Bloco é a inicial ("0")
        PosicaoDoBloco = 0

        TimerJogo.Enabled = True
        TimerMovimento.Enabled = True
        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(1)
        cmdNovoJogoSTATUS = 1
        
        'Desabilita os menus selecionadores de Nível de Jogo
        MenuNivel ("Desabilitar")
        
    ElseIf cmdNovoJogo.Caption = cmdNovoJogoTEXTO(1) Then
        'Pausa o Jogo
    
        TimerJogo.Enabled = False
        TimerMovimento.Enabled = False
        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(2)
        cmdNovoJogoSTATUS = 2
        
    Else
        'Reinicia o Jogo, após Pausar
        
        TimerJogo.Enabled = True
        TimerMovimento.Enabled = True
        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(1)
        cmdNovoJogoSTATUS = 1
        
    End If
  
    pctJogo.SetFocus
    
End Function

Function PrepararJogo()
'Prepara o Jogo para ser iniciado

    Dim Contador1, Contador2 As Integer
    
    LimpaTela
    
    'Armazena "False" para todas os blocos de "pctJogo",
    'retornando-os a sua posição inicial
    Contador1 = 0
    
    Do While Contador1 < 180
       
        BlocoEmJogo(Contador1) = False
    
        Contador1 = Contador1 + 1
    
    Loop
    
    'Armazena "999" para todas as posições da matriz "EstiloBlocoEmJogo"
    Contador1 = 0
    
    Do While Contador1 < 180
    
        EstiloBlocoEmJogo(Contador1) = 999
    
        Contador1 = Contador1 + 1
    
    Loop
    
    'Armazena "999" em todas as posições da matriz "Jogo"
    Contador1 = 1
    Contador2 = 1

    Do While Contador1 <= 18
    
        Do While Contador2 <= 10
        
            Jogo(Contador1, Contador2) = 999

            Contador2 = Contador2 + 1
        
        Loop
        
        Contador2 = 1
        
        Contador1 = Contador1 + 1
    
    Loop
    
End Function
