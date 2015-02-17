VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJogo 
   Caption         =   "Down Blocks"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6030
   Icon            =   "frmJogo.frx":0000
   LinkTopic       =   "frmJogo"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerMovimento 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   4320
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
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.PictureBox pctProx 
      BackColor       =   &H00808000&
      Height          =   900
      Left            =   4070
      ScaleHeight     =   840
      ScaleWidth      =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1860
      Begin VB.PictureBox pctProxBloc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   191
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctProxBloc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   190
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctProxBloc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   189
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctProxBloc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   188
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.PictureBox pctJogo 
      BackColor       =   &H00808000&
      Height          =   6810
      Left            =   120
      ScaleHeight     =   6750
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   120
      Width           =   3810
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   179
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   187
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   178
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   186
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   177
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   185
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   176
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   184
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   175
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   183
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   174
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   182
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   173
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   181
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   172
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   180
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   171
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   179
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   170
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   178
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   169
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   177
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   168
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   176
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   167
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   175
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   166
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   174
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   165
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   173
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   164
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   172
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   163
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   171
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   162
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   170
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   161
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   169
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   160
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   168
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   159
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   167
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   158
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   166
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   157
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   165
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   156
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   164
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   155
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   163
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   154
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   162
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   153
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   161
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   152
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   160
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   151
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   159
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   150
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   158
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   149
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   157
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   148
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   156
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   147
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   155
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   146
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   154
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   145
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   153
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   144
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   152
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   143
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   151
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   142
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   150
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   141
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   149
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   140
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   148
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   139
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   147
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   138
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   146
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   137
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   136
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   144
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   135
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   143
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   134
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   142
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   133
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   141
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   132
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   140
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   131
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   139
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   130
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   138
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   129
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   137
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   128
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   127
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   135
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   126
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   134
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   125
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   124
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   132
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   123
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   131
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   122
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   130
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   121
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   129
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   120
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   128
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   119
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   127
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   118
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   126
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   117
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   116
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   124
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   115
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   123
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   114
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   122
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   113
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   121
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   112
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   111
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   119
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   110
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   118
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   109
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   117
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   108
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   116
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   107
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   115
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   106
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   114
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   105
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   113
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   104
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   112
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   103
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   111
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   102
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   110
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   101
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   109
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   100
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   108
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   99
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   107
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   98
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   106
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   97
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   105
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   96
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   104
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   95
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   103
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   94
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   93
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   101
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   92
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   100
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   91
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   99
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   90
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   98
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   89
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   97
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   88
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   96
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   87
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   95
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   86
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   94
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   85
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   93
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   84
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   83
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   91
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   82
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   90
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   81
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   89
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   80
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   79
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   78
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   77
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   85
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   76
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   75
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   83
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   74
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   82
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   73
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   81
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   72
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   80
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   71
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   79
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   70
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   78
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   69
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   77
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   68
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   67
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   75
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   66
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   65
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   64
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   72
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   63
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   62
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   70
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   61
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   60
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   68
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   59
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   58
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   57
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   56
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   55
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   54
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   53
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   52
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   51
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   50
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   49
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   48
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   47
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   46
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   45
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   44
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   43
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   42
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   41
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   40
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   39
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   38
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   37
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   36
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   35
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   34
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   33
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   32
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   31
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   30
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   29
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   28
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   27
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   26
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   25
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   24
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   23
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   22
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   21
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   20
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctBloco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu mnuJogo 
      Caption         =   "Jogo"
      Begin VB.Menu mnuNovoJogo 
         Caption         =   "Novo Jogo"
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
Dim Nivel As Integer ' Indica o Nível (velocidade de descida
                     'dos Blocos) do jogo
Dim TeclaPressionada As Integer 'Indica a tecla que foi pressionada
Dim cmdNovoJogoSTATUS As Integer 'Indica qual a posição da va-
                                 'riável "cmdNovoJogoTEXTO"
                                 'está sendo usada no momento
                                 'como rótulo em "cmdNovoJogo"

Private Sub pctJogo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    TeclaPressionada = KeyCode

End Sub

Private Sub TimerMovimento_Timer()

    Select Case TeclaPressionada
    
        Case 38 'Seta para cima
            If pctBloco(0).Top > 0 Then
            
                pctBloco(0).Top = pctBloco(0).Top - 375
            
            End If
            TeclaPressionada = 0
        
        Case 40 'Seta para baixo
            If pctBloco(0).Top < 6000 Then
                    
                If mnuSons.Checked = True Then
                
                    TocarSom App.Path & "\Sons\Descendo.wav", SND_ASYNC
                
                End If
                
                pctBloco(0).Top = pctBloco(0).Top + 750
                               
            End If
            TeclaPressionada = 0
        
        Case 37 'Seta para o lado esquerdo
            If pctBloco(0).Left > 0 Then
            
                pctBloco(0).Left = pctBloco(0).Left - 375
         
            End If
            TeclaPressionada = 0
        
        Case 39 'Seta para cima direito
            If pctBloco(0).Left < 3375 Then
            
                pctBloco(0).Left = pctBloco(0).Left + 375
            
            End If
            TeclaPressionada = 0
        
    End Select

End Sub

Private Sub TimerJogo_Timer()
  
    If pctBloco(0).Top < 6375 Then

        If mnuSons.Checked = True Then

            TocarSom App.Path & "\Sons\Descendo.wav", SND_ASYNC
        
        End If
        
        pctBloco(0).Top = pctBloco(0).Top + 375
        
    End If

End Sub



Private Sub cmdNovoJogo_Click()
'Primeiramente, limpa a tela de Jogo e, depois, inicia um
'novo Jogo, habilitando o Timer ("TimerJogo")
    
    If cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0) Then
        
        'Inicia o Jogo
        
        'Gera o primeiro bloco...
        GerarBlocos
        
        TimerJogo.Enabled = True
        TimerMovimento.Enabled = True
        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(1)
        cmdNovoJogoSTATUS = 1
        
        With pctBloco(0)
            .Visible = True
            .Top = 0
            .Left = 1125
            .Picture = ImageListBlocos.ListImages(1).Picture
        End With
        
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
    
End Sub

Private Sub Form_Load()
' Logo ao iniciar, carrega o ImageList com o Estilo de Bloco
'"Clássico", e define o nível inicial como 1 (1000 milissegundos)
    CarregarImageList (App.Path & "\Blocos\Clássico\")
    Nivel = 1000
    
    'Indica o nível do Jogo (Inicialmente "0" - 1000 milissegundos)
    TimerJogo.Interval = Nivel
    
    'Carrega os rótulos de "cmdNovoJogo" na variável
    '"cmdNovoJogoTEXTO"
    cmdNovoJogoTEXTO(0) = "Iniciar"
    cmdNovoJogoTEXTO(1) = "Pausar"
    cmdNovoJogoTEXTO(2) = "Continuar"
    
    'Indica o rótulo utilizado em "cmdNovoJogo" (no caso a
    'posição "0" da variável "cmdNovoJogoTEXTO"
    cmdNovoJogoSTATUS = 0

End Sub

'===========================================================
'COMANDOS REFERENTES AOS MENUS                             =
'===========================================================
Private Sub mnuNovoJogo_Click()
'Primeiramente, limpa a tela de Jogo e, depois, inicia um
'novo Jogo, habilitando o Timer ("TimerJogo")
    TimerJogo.Enabled = True
    
End Sub

Private Sub mnuSair_Click()
'Finaliza o jogo
    End

End Sub

Private Sub mnuSobre_Click()
'Exibe "frmSobre"
    frmSobre.Show (vbModal)

End Sub

Private Sub mnuEst_Click(Index As Integer)
' Seleciona o Estilo do Bloco, conforme a seleção do usuário
'(indicada pela variável "Index")
    
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

End Sub

Private Sub mnuNivel_Click(Index As Integer)
' Seleciona o Nível do Jogo (velocidade de descida dos Blocos),
'conforme a seleção do usuário (indicada pela variável "Index")
    
    'Tira o "Checked" dos menus de Níveis
    Uncheck ("Nível")
    
    'Através da variável "Index", seleciona o Nível do Jogo
    Select Case Index
    
        Case 0
            Nivel = 1000
            lblNivel = 0
            TimerJogo.Interval = Nivel
            mnuNivel(0).Checked = True
            
        Case 1
            Nivel = 880
            lblNivel = 1
            TimerJogo.Interval = Nivel
            mnuNivel(1).Checked = True
            
        Case 2
            Nivel = 760
            lblNivel = 2
            TimerJogo.Interval = Nivel
            mnuNivel(2).Checked = True
            
        Case 3
            Nivel = 640
            lblNivel = 3
            TimerJogo.Interval = Nivel
            mnuNivel(3).Checked = True
            
        Case 4
            Nivel = 520
            lblNivel = 4
            TimerJogo.Interval = Nivel
            mnuNivel(4).Checked = True
            
        Case 5
            Nivel = 400
            lblNivel = 5
            TimerJogo.Interval = Nivel
            mnuNivel(5).Checked = True
            
        Case 6
            Nivel = 280
            lblNivel = 6
            TimerJogo.Interval = Nivel
            mnuNivel(6).Checked = True
            
        Case 7
            Nivel = 160
            lblNivel = 7
            TimerJogo.Interval = Nivel
            mnuNivel(7).Checked = True
            
        Case 8
            Nivel = 40
            lblNivel = 8
            TimerJogo.Interval = Nivel
            mnuNivel(8).Checked = True
            
        Case 9
            Nivel = 1
            lblNivel = 9
            TimerJogo.Interval = Nivel
            mnuNivel(9).Checked = True
            
    End Select

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
