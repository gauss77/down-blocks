VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJogo 
   Caption         =   "Down Blocks"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9690
   Icon            =   "frmJogo.frx":0000
   LinkTopic       =   "frmJogo"
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "T"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   58
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   57
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "| |"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   56
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   55
      Top             =   6000
      Width           =   375
   End
   Begin VB.Frame fraNovoRecorde 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1440
      TabIndex        =   47
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtNomeJogador 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   48
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdGravarRecorde 
         Caption         =   "Ok"
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   840
         Width           =   555
      End
      Begin VB.Line LineFraNovoRecorde 
         Index           =   3
         X1              =   3360
         X2              =   3360
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Line LineFraNovoRecorde 
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Line LineFraNovoRecorde 
         Index           =   1
         X1              =   0
         X2              =   3360
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line LineFraNovoRecorde 
         Index           =   0
         X1              =   0
         X2              =   3360
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblTXTNovoRecordeNome 
         BackColor       =   &H00808080&
         Caption         =   "Seu Nome:"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTXTNovoRecorde 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Você obteve um Novo recorde!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Frame fraRecordes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   1440
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdFecharRecordes 
         Appearance      =   0  'Flat
         Caption         =   "Fechar"
         Height          =   315
         Left            =   2280
         TabIndex        =   44
         Top             =   3240
         Width           =   1020
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   34
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   33
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   32
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   31
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   30
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   29
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblJogador 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.Line LineFraRecordes 
         Index           =   3
         X1              =   3360
         X2              =   3360
         Y1              =   0
         Y2              =   3600
      End
      Begin VB.Line LineFraRecordes 
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3600
      End
      Begin VB.Line LineFraRecordes 
         Index           =   1
         X1              =   0
         X2              =   3360
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line LineFraRecordes 
         Index           =   0
         X1              =   0
         X2              =   3360
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   43
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   42
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   41
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   40
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   39
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   38
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   36
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   35
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "10."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "9."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "8."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "7."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "6."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblPosicao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblPontuacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblTXTRecordes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recordes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Timer TimerMovimento 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   5400
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Depuração do Sistema "
      Height          =   6855
      Left            =   6120
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdConfigINI 
         Caption         =   "Config. Padrão"
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton cmdDebugScore 
         Caption         =   "Limpar Recordes"
         Height          =   375
         Left            =   1920
         TabIndex        =   53
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtDebug 
         Appearance      =   0  'Flat
         Height          =   3615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton cmdDebugEsconder 
         Caption         =   "Finalizar Depuração"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label lblDebugMatrizTexto 
         Caption         =   "Matriz ""Jogo"""
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3255
      End
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
      Top             =   5400
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
      Top             =   5400
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
      Top             =   6480
      Width           =   1575
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
   Begin VB.CommandButton cmdTESTE 
      Caption         =   "TESTAR"
      Height          =   255
      Left            =   4080
      TabIndex        =   52
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblTXTRecorde 
      Alignment       =   2  'Center
      Caption         =   "Recorde"
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
      TabIndex        =   46
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblRecorde 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   45
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   4080
      X2              =   5880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4440
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
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblPontos 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   5880
      Y1              =   3840
      Y2              =   3840
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
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   5880
      Y1              =   2640
      Y2              =   2640
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
      Begin VB.Menu mnuEspaco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEstiloTexto 
         Caption         =   "Estilo dos Blocos"
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
      Begin VB.Menu mnuEspaco3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNivelTexto 
         Caption         =   "Nível de Início"
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
      Begin VB.Menu mnuEspaco4 
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
      Begin VB.Menu mnuEspaco5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordes 
         Caption         =   "Ver Recordes"
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
      Begin VB.Menu mnuEspaco6 
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
    
    'VARIÁVEIS de uso nos eventos de "TimerMovimento"
    Dim PosicaoBlocoXY As Posicao
    Dim ContadorTimerMovimento As Integer
    Dim ColidiuEsquerdaDireita As Boolean
    Dim Blocos_a_VerificarTEMP As PosicoesParaIndices
    Dim Blocos_a_Verificar(4) As Integer
    Dim Linha_Completa As Integer
    Dim CriarProxBloco As Boolean
    Dim Linha_da_colisao As Integer
    Dim BlocosLinhaIguais As Boolean
    Dim EstiloPrimeiroBloco As Integer
    Dim PosicaoRecorde As Integer
    Dim MoverBlocoParaBaixo As Boolean ' Indica se o Bloco deve
                                       '("MoverBlocoParaBaixo = True")
                                       'ser movido para baixo ou não
                                       '("MoverBlocoParaBaixo = False")






Private Sub cmdTESTE_Click()

'GerarBlocos "pctJogo", 1, 1

'MoverBloco 40, 1, 999, 999, 999, 1, 0
'ExibirValJogo
'Command1.Enabled = True

If mnuMusica.Enabled = True Then

    TocarMusica Tocar, App.Path & "\Musicas\Clotho.mid"

End If

'Armazena o tempo inicial da música
            TempoMusica = 100
            
            MusicaExecutando = True
            
TocarMusica Tempo, , TempoMusica

Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True

Exit Sub

Recordes "Armazenar", 6, "Carlos", 40
Recordes "CarregarVariavel"
Recordes "Exibir"

fraRecordes.Visible = True
   
   

End Sub
Private Sub Command1_Click()
TocarMusica Resumir
Command2.Enabled = True
Command1.Enabled = False
End Sub
Private Sub Command2_Click()

    TocarMusica Pausar

Command2.Enabled = False
Command1.Enabled = True
End Sub
Private Sub Command3_Click()
TocarMusica Parar
MusicaExecutando = False

Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
End Sub
Private Sub Command4_Click()
TocarMusica Tempo, , 100

End Sub





Private Sub cmdNovoJogo_Click()
    
    'Inicia o Jogo
    IniciarJogo

End Sub

Private Sub cmdFecharRecordes_Click()
'Esconde "fraRecordes"

    fraRecordes.Visible = False
    
    'Tira possíveis negritos e cores vermelhas dos textos
    ContadorFecharRecordes = 0
    
    lblPosicao(PosicaoRecorde).ForeColor = vbBlack
        
    lblJogador(PosicaoRecorde).ForeColor = vbBlack
    lblJogador(PosicaoRecorde).FontBold = False
        
    lblPontuacao(PosicaoRecorde).ForeColor = vbBlack
    lblPontuacao(PosicaoRecorde).FontBold = False
    
    'Desbloqueia os componentes de "frmJogo"
    Componentes ("Desbloquear")

End Sub

Private Sub cmdGravarRecorde_Click()
' Esconde "fraNovoRecorde", armazena o novo recorde,
'exibe "fraRecordes" já com o novo recorde armazenado
'e grifado em negrito e vermelho

    ' Verifica se o nome do jogador não contém o caractere "%"
    '(de uso especial da função "Recordes"
    If InStr(1, txtNomeJogador.Text, "%") = 0 Then

        fraNovoRecorde.Visible = False
        
        Recordes "Armazenar", PosicaoRecorde, txtNomeJogador.Text, CLng(lblPontos.Caption)
        Recordes "CarregarVariavel"
        Recordes "Exibir"
        
        ' Como os índices controles estão sempre UMA POSIÇÃO a menos,
        'subtraí-se "1" de "PosicaoRecorde" (variável que será utilizada
        'como base - servindo de índice - ao negritar os textos dos
        'controles referentes à posição do recorde do jogador atual)
        PosicaoRecorde = PosicaoRecorde - 1
        
        'Indica em negrito e cor vermelha a pontuação do jogador
        lblPosicao(PosicaoRecorde).ForeColor = vbRed
        
        lblJogador(PosicaoRecorde).ForeColor = vbRed
        lblJogador(PosicaoRecorde).FontBold = True
        
        lblPontuacao(PosicaoRecorde).ForeColor = vbRed
        lblPontuacao(PosicaoRecorde).FontBold = True
        
        fraRecordes.Visible = True
        
        'Desbloqueia os componentes de "frmJogo"
        Componentes ("Desbloquear")
        
        'Limpa o conteúdo da Caixa de Texto
        txtNomeJogador.Text = ""
        
        lblRecorde.Caption = lblPontuacao(0).Caption
    
    Else
    
        MsgBox MsgBoxGravarRecorde, vbOKOnly + vbExclamation, "Down Blocks"
       
        txtNomeJogador.SetFocus
        
    End If
    
End Sub

Private Sub pctJogo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Armazena a tecla pressionada pelo usuário
    TeclaPressionada = KeyCode

End Sub

Private Sub TimerMovimento_Timer()
'Realiza os movimentos do Bloco

'On Error GoTo Erro

    ' Verifica se deve haver movimento para esquerda, direita,
    'mudança de posição do Bloco ou aceleração do movimento
    Select Case TeclaPressionada
    
        Case 38 'Seta para cima
            'Muda as posições do Bloco
            
            GoTo MudarPosicaoBloco
        
        Case 40 'Seta para baixo
            'Acelera a descida do Bloco
            
            ' Para acelerar a descida do Bloco, indica-se "TRUE"
            'como valor para a variável "MoverBlocoParaBaixo",
            'forçando o jogo a executar a descida do Bloco a cada
            'execução de "TimerMovimento" (realizada a cada 1
            'milissegundo)
            MoverBlocoParaBaixo = True

        Case 37 'Seta para o lado esquerdo
            'Move o Bloco uma posição para a esquerda
            
            Blocos_a_VerificarTEMP = ArmazenarPosicoes(37, TipoBlocoEmJogo, PosicaoDoBloco)
            Blocos_a_Verificar(0) = Blocos_a_VerificarTEMP.idx1
            Blocos_a_Verificar(1) = Blocos_a_VerificarTEMP.idx2
            Blocos_a_Verificar(2) = Blocos_a_VerificarTEMP.idx3
            Blocos_a_Verificar(3) = Blocos_a_VerificarTEMP.idx4

            GoTo MoverEsquerdaDireita

        Case 39 'Seta para o lado direito
            'Move o Bloco uma posição para a direita
            
            Blocos_a_VerificarTEMP = ArmazenarPosicoes(39, TipoBlocoEmJogo, PosicaoDoBloco)
            Blocos_a_Verificar(0) = Blocos_a_VerificarTEMP.idx1
            Blocos_a_Verificar(1) = Blocos_a_VerificarTEMP.idx2
            Blocos_a_Verificar(2) = Blocos_a_VerificarTEMP.idx3
            Blocos_a_Verificar(3) = Blocos_a_VerificarTEMP.idx4

            GoTo MoverEsquerdaDireita

    End Select

    GoTo MoverParaBaixo

    Exit Sub
    
    
    
    
    
'CÓDIGOS PARA MUDA A POSIÇÃO DO BLOCO
MudarPosicaoBloco:
    
    'Primeiramente, deve-se verificar o Tipo do Bloco em Jogo
    'Verifica qual o tipo de Bloco
    Select Case TipoBlocoEmJogo
            
        Case 0
        'Blocos do tipo ****
                    
            'Verifica a posição atual do bloco
            Select Case PosicaoDoBloco
                    
                Case 0
                ' Sendo a posição 0, o Bloco mover-se-á
                'para a posição 1
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           XX
                    '           1234
                    '            XXX
                    '            XX
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica-se a posição do 1° bloco do conjunto:
                    PosicaoBlocoXY = LocalizarBlocoNaMatriz(IndicesEmJogo(1))
                    
                    If PosicaoBlocoXY.PosicaoY = 1 Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Acima do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. Abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
    
                    '4. Abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '5. Abaixo do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. 2 blocos abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '7. 2 blocos abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda...
                    MoverBloco 37, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '3. ...novamente para baixo...
                    MoverBloco 40, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '4. ...e, por fim, para a esquerda
                    MoverBloco 37, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Movimentos do bloco 3 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 1
                        
                Case 1
                ' Sendo a posição 1, o Bloco mover-se-á
                'para a posição 0
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           X1
                    '           X2XX
                    '            3XX
                    '            4X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                                       
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À esquerda do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À esquerda do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. À direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
    
                    '4. 2 blocos à direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(3), 39, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '5. À direita do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. 2 blocos à direita do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '7. À direita do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para baixo
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 3 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima
                    MoverBloco 38, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima...
                    MoverBloco 38, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '3. ...novamente para a direita...
                    MoverBloco 39, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '4. ...e, por fim, para cima
                    MoverBloco 38, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 0
                        
            End Select
                    
        Case 1
        'Blocos do tipo ***
                    
            'Verifica a posição atual do bloco
            Select Case PosicaoDoBloco
                    
                Case 0
                ' Sendo a posição 0, o Bloco mover-se-á
                'para a posição 1
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           XX
                    '           123
                    '            XX
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica-se a posição do 1° bloco do conjunto:
                    PosicaoBlocoXY = LocalizarBlocoNaMatriz(IndicesEmJogo(1))
                    
                    If PosicaoBlocoXY.PosicaoY = 1 Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Acima do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. Abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. Abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 3 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 1
                     
                Case 1
                ' Sendo a posição 1, o Bloco mover-se-á
                'para a posição 0
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           X1
                    '           X2X
                    '            3X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À esquerda do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À esquerda do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. À direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
    
                    '4. À direita do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para baixo
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 3 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima
                    MoverBloco 38, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 0
                        
            End Select
                
        Case 2
        '                *
        'Blocos do tipo ***
                    
            'Verifica a posição atual do bloco
            Select Case PosicaoDoBloco
                    
                Case 0
                ' Sendo a posição 0, o Bloco mover-se-á
                'para a posição 1
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           X1X
                    '           234
                    '           XX
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)

                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À esquerda do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À direita do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. Abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. Abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para baixo
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 1
                        
                Case 1
                ' Sendo a posição 1, o Bloco mover-se-á
                'para a posição 2
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           X4
                    '           13X
                    '           X2X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)

                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À esquerda do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À direita do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. À esquerda do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. À direita do Bloco2
                    If DetectarColisao(IndicesEmJogo(2), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima
                    MoverBloco 38, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a baixo
                    MoverBloco 40, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 2
                            
                Case 2
                ' Sendo a posição 2, o Bloco mover-se-á
                'para a posição 3
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '            XX
                    '           432
                    '           X1X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)

                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Abaixo do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. Abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para a cima...
                    MoverBloco 38, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a cima
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 3
                        
                Case 3
                ' Sendo a posição 3, o Bloco mover-se-á
                'para a posição 0
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           X2X
                    '           X31
                    '            4X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)

                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À direita do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À esquerda do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. À direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. À esquerda do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para baixo
                    MoverBloco 40, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima
                    MoverBloco 38, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 0
                        
            End Select
                
        Case 3
        '                 *
        'Blocos do tipo ***
        
            'Verifica a posição atual do bloco
            Select Case PosicaoDoBloco
                    
                Case 0
                ' Sendo a posição 0, o Bloco mover-se-á
                'para a posição 1
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '            XX
                    '           XX1
                    '           234
                    '           XX
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica-se a posição do 1° bloco do conjunto:
                    PosicaoBlocoXY = LocalizarBlocoNaMatriz(IndicesEmJogo(1))
                    
                    If PosicaoBlocoXY.PosicaoY = 1 Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Acima do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. Abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. Acima do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If

                    '5. 2 blocos acima do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 38, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. Abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '0. Move-o duas vezes para a esquerda
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 1
                        
                Case 1
                ' Sendo a posição 1, o Bloco mover-se-á
                'para a posição 2
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '          X14
                    '          XX3X
                    '           X2X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À esquerda do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. À esquerda do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. À direita do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If

                    '5. 2 blocos à esquerda do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 37, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. À esquerda Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '0. Move-o duas vezes para baixo
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para direita...
                    MoverBloco 39, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a cima
                    MoverBloco 38, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para baixo
                    MoverBloco 40, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 2
                            
                Case 2
                ' Sendo a posição 2, o Bloco mover-se-á
                'para a posição 3
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           XX
                    '          432
                    '          1XX
                    '          XX
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Abaixo do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. Abaixo do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. Acima do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If

                    '5. 2 blocos abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. Abaixo Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '0. Move-o duas vezes para a direita
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 37, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 3
     
                Case 3
                ' Sendo a posição 3, o Bloco mover-se-á
                'para a posição 0
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '         X2X
                    '         X3XX
                    '          41X
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. À direita do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. À esquerda do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '3. À direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. À esquerda do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 37) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If

                    '5. 2 blocos à direita do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. À direita Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 39) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 1 do conjunto
                    '0. Move-o duas vezes para cima
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 2 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para baixo
                    MoverBloco 40, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 4 do conjunto
                    '1. Move-o primeiramente para a direita...
                    MoverBloco 39, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima
                    MoverBloco 38, IndicesEmJogo(4), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 0
                        
            End Select
                
        Case 4
        '                **
        'Blocos do tipo **
                    
            'Verifica a posição atual do bloco
            Select Case PosicaoDoBloco
                    
                Case 0
                ' Sendo a posição 0, o Bloco mover-se-á
                'para a posição 1
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '           XXX
                    '           X12
                    '           34
                    '           XX
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)
                    
                    'Verifica-se a posição do 1° bloco do conjunto:
                    PosicaoBlocoXY = LocalizarBlocoNaMatriz(IndicesEmJogo(1))
                    
                    If PosicaoBlocoXY.PosicaoY = 1 Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Acima do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If

                    '3. Acima do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. 2 blocos acima do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 38, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '5. Abaixo do Bloco 3
                    If DetectarColisao(IndicesEmJogo(3), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '6. Abaixo do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 3 do conjunto
                    '1. Move-o primeiramente para baixo...
                    MoverBloco 40, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a direita
                    MoverBloco 39, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para a esquerda...
                    MoverBloco 37, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a baixo
                    MoverBloco 40, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 2 do conjunto
                    '0. Move-o duas vezes para a esquerda...
                    MoverBloco 37, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    MoverBloco 37, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 1

                Case 1
                ' Sendo a posição 1, o Bloco mover-se-á
                'para a posição 0
                
                    'Neste caso, deve-se verificar as posições:
                    'Obs.: O X indica essas posições; os números
                    'são as referências aos índices do Bloco no
                    'Jogo
                    '
                    '          XX
                    '          2XX
                    '          14
                    '          X3
                    '
                    'Sendo as posições verificadas e concluídas
                    'como devidamente livres, move-se o Bloco.
                    'ATENÇÃO: Antes de iniciar-se essa verificação,
                    'há a análise da linha atual do 1° bloco do
                    'conjunto. Este deve estar no mínimo na linha 2
                    'da matriz (indicando, assim, a possibilidade de
                    'movimentação)

                    'Verifica se há colisão:
                    ' Obs.: Caso haja colisão, abandonam-se os
                    'procedimentos de troca de Posição (em 360°)
                    'do Bloco e passa-se para a Movimentação do
                    'Bloco para baixo)
                    
                    '1. Abaixo do Bloco 1
                    If DetectarColisao(IndicesEmJogo(1), 40) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '2. Acima do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If

                    '3. Acima do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 38) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '4. 2 blocos acima do Bloco 4
                    If DetectarColisao(IndicesEmJogo(4), 38, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                    
                    '5. 2 blocos à direita do Bloco 2
                    If DetectarColisao(IndicesEmJogo(2), 39, 2) = True Then
                    
                        GoTo MoverParaBaixo
                    
                    End If
                                        
                    ' Não havendo nenhuma colisão, move o Bloco para
                    'a nova Posição
                    
                    'Movimentos do bloco 2 do conjunto
                    '0. Move-o duas vezes para a direita...
                    MoverBloco 39, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    MoverBloco 39, IndicesEmJogo(2), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 1 do conjunto
                    '1. Move-o primeiramente para cima...
                    MoverBloco 38, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para a esquerda
                    MoverBloco 39, IndicesEmJogo(1), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    
                    'Movimentos do bloco 3 do conjunto
                    '1. Move-o primeiramente para esquerda...
                    MoverBloco 37, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco
                    '2. ...depois para cima
                    MoverBloco 38, IndicesEmJogo(3), 999, 999, 999, TipoBlocoEmJogo, PosicaoDoBloco

                    'Indica a nova posição que o Bloco receberá
                    PosicaoDoBloco = 0
                        
            End Select
                    
        Case 5
        '               **
        'Blocos do tipo **
                    
            ' Neste tipo de bloco, nenhum movimento em 360° ocorre.
            'Assim sendo, indica-se ao Jogo para que este prossiga
            'com os movimentos de descida
                    
    End Select

    GoTo MoverParaBaixo
    
    
    
    
    
'CÓDIGOS PARA MOVIMENTAR O BLOCO PARA A ESQUERDA OU DIREITA
MoverEsquerdaDireita:
    
    'Analisa as posições solicitadas do Bloco em Jogo
    ' Obs.: A informação "999" armazenada em "Blocos_a_Verificar"
    'indica que aquela posição não deve ser verificada
    ContadorTimerMovimento = 0
    
    ColidiuEsquerdaDireita = False
    
    Do While ContadorTimerMovimento <= 3
    
        If Blocos_a_Verificar(ContadorTimerMovimento) <> 999 Then
    
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
                    If DetectarColisao(IndicesEmJogo(Blocos_a_Verificar(ContadorTimerMovimento)), 37) = True Then

                         ColidiuEsquerdaDireita = True
                
                        Exit Do
                
                    End If
                
                Case 39 'Lado direito
                    'Verifica se há colisão
                    If DetectarColisao(IndicesEmJogo(Blocos_a_Verificar(ContadorTimerMovimento)), 39) = True Then

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
      IndicesEmJogo(3), IndicesEmJogo(4), TipoBlocoEmJogo, PosicaoDoBloco
         
    
    
    
    
'CÓDIGOS PARA A REALIZAÇÃO DA DESCIDA DO BLOCO
MoverParaBaixo:





ExibirValJogo






    TeclaPressionada = 0

    'Verifica se o Bloco deve descer para baixo
    If MoverBlocoParaBaixo = True Then
    
            MoverBlocoParaBaixo = False
    
            CriarProxBloco = False
            
            Blocos_a_VerificarTEMP = ArmazenarPosicoes(40, TipoBlocoEmJogo, PosicaoDoBloco)
            Blocos_a_Verificar(0) = Blocos_a_VerificarTEMP.idx1
            Blocos_a_Verificar(1) = Blocos_a_VerificarTEMP.idx2
            Blocos_a_Verificar(2) = Blocos_a_VerificarTEMP.idx3
            Blocos_a_Verificar(3) = Blocos_a_VerificarTEMP.idx4
            
            'Verifica as posições de blocos solicitadas
            ' Obs.: A informação "999" armazenada em "Blocos_a_Verificar"
            'indica que aquela posição não deve ser verificada
            ContadorTimerMovimento = 0
            
            Do While ContadorTimerMovimento <= 3
            
                If Blocos_a_Verificar(ContadorTimerMovimento) <> 999 Then
                
                    'Verifica se há colisão
                    If DetectarColisao(IndicesEmJogo(Blocos_a_Verificar(ContadorTimerMovimento)), 40) = True Then
        
                        CriarProxBloco = True
                        
                        Exit Do
                        
                    End If
                
                End If
            
                ContadorTimerMovimento = ContadorTimerMovimento + 1
            
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

            'PONTUAÇÃO DO JOGO====================================================
                    ' Marca-se pontos caso uma linha do jogo esteja completa.
                    ' Se a linha completa contiver blocos de mesmo estilo,
                    'adiciona-se "300" à pontuação do jogador. Caso contrário,
                    'adiciona-se "50" à pontuação do jogador.
                    ' Troca-se de nível a cada 1500 pontos e termina-se o jogo
                    'com 15000 pontos
        
                    'Verifica se o Estilo dos Blocos desta linha são os mesmos
                  
                    'Armazena o estilo do primeiro bloco
                    EstiloPrimeiroBloco = EstiloBlocoEmJogo(Jogo(Linha_Completa, 1))
                    
                    ContadorTimerMovimento = 2
                    BlocosLinhaIguais = True
                    
                    Do While ContadorTimerMovimento <= 10
                    
                        If EstiloBlocoEmJogo(Jogo(Linha_Completa, _
                         ContadorTimerMovimento)) <> EstiloPrimeiroBloco _
                         Then BlocosLinhaIguais = False
                    
                        ContadorTimerMovimento = ContadorTimerMovimento + 1
                    
                    Loop
                    
                    If BlocosLinhaIguais = True Then
                    
                        'Aumenta-se em 300 a pontuação
                        lblPontos.Caption = CStr(CLng(lblPontos.Caption) + 300)
                    
                    Else
                    
                        'Aumenta-se em 50 a pontuação
                        lblPontos.Caption = CStr(CLng(lblPontos.Caption) + 50)
                    
                    End If
                    
                    'Verifica se há troca de nível
                    ContadorTimerMovimento = 1
                    
                    Do While ContadorTimerMovimento <= 9

                        If CLng(lblPontos.Caption) = (CLng(ContadorTimerMovimento) * 1500) Then
                        'Aumenta um nível

                            If (CLng(lblNivel.Caption) + 1) <= 9 _
                              Then SelecionarNivel _
                              (CLng(lblNivel.Caption) + 1), True
    
                        ElseIf CLng(lblPontos.Caption) >= 15000 Then 'Salvou o Jogo!
'SALVOU O JOGO=========================================
                            MsgBox "Salvou o Jogo!!!"
                    
                        End If
                        
                        ContadorTimerMovimento = ContadorTimerMovimento + 1
                        
                    Loop
            'FIM DOS CÓDIGOS DE PONTUAÇÃO DO JOGO=================================
                    
                    ContadorTimerMovimento = 1
                    
                    Do While ContadorTimerMovimento <= 10
                    
                        'Marca o bloco como não mais em uso
                        BlocoEmJogo(Jogo(Linha_Completa, ContadorTimerMovimento)) = False
                        
                        'Esconde-se o bloco, retornando-o à posição inicial
                        With imgBloco(Jogo(Linha_Completa, ContadorTimerMovimento))
                            .Top = 0
                            .Left = 0
                            .Visible = False
                            .Picture = Nothing
                        End With
                        
                        ' Armazena "999" em "EstiloBlocoEmJogo", indicando que não há
                        'figura neste bloco (o bloco não está em uso)
                        EstiloBlocoEmJogo(Jogo(Linha_Completa, ContadorTimerMovimento)) = 999
                        
                        'Armazena 999 na posição deste bloco na matriz "Jogo"
                        Jogo(Linha_Completa, ContadorTimerMovimento) = 999
                    
                        ContadorTimerMovimento = ContadorTimerMovimento + 1
                    
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
                
        'AFERIÇÃO DO FIM DE JOGO==============================================
                
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
                
                ContadorTimerMovimento = 4
                
                Do While ContadorTimerMovimento <= 6
'FIM DE JOGO ===================================
                    If Jogo(Linha_da_colisao, ContadorTimerMovimento) <> 999 Then
                    'Fim de jogo
                        
MsgBox "Fim de jogo..."
                        
                        TimerJogo.Enabled = False
                        TimerMovimento.Enabled = False
                        
                        'O menu de término de jogo é desabilitado
                        mnuFinalizar.Enabled = False
 
                        'Habilita os menus selecionadores de Nível de Jogo
                        MenuNivel ("Habilitar")
                        
                        'Coloca o rótulo inicial em "cmdNovoJogo"
                        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
                         
                        'Se o jogador não fez pontos, simplesmente sai do Jogo
                        If lblPontos.Caption = "0" Then Exit Sub
 
                        'Verifica se a pontuação do jogador bate algum recorde
                        ContadorTimerMovimento = 1 ' O contador pode ser zerado
                                                   'sem problemas, uma vez que
                                                   'saindo-se deste procedimento
                                                   'o jogo será finalizado
                        
                        Do While ContadorTimerMovimento <= 10

                            ' Verifica se a pontuação na posição atual é igual
                            'à pontuação obtida pelo jogador
                            If RecordesJogo(ContadorTimerMovimento, 2) = CLng(lblPontos.Caption) Then
                            'Sendo igual, verifica se há outro igual a este valor logo
                            'abaixo desta posição (Isso assegura que, se a pontuação
                            'deste jogador seja igual a de outros, o recorde deste
                            'seja acrescido após os outros iguais já existentes
                                
                                ' Verifica-se se "ContadorTimerMovimento + 1" pois
                                'há apenas 10 posições de recordes para serem utilizadas
                                If (ContadorTimerMovimento + 1) <= 10 Then
                                
                                    If RecordesJogo((ContadorTimerMovimento + 1), 2) = CLng(lblPontos.Caption) Then
                                    'Sendo o próximo igual, volta para a realização do DO
                                    
                                        GoTo Continuar_DO
                                    
                                    Else
                                    ' Caso não haja mais nenhum igual, isso indica que este recorde
                                    'pode ser armazenado UMA posição à frente desta (pois esta posição
                                    'é de IGUAL valor à pontuação do jogador. Assim, em caso de empate,
                                    'o valor que já estiver armazenado permanece anteriormente à um
                                    'novo valor que seja igual.
                                        
                                    PosicaoRecorde = (ContadorTimerMovimento + 1)
                                            
                                    fraNovoRecorde.Visible = True
                                    txtNomeJogador.SetFocus
                                            
                                    'Bloqueia os componentes de "frmJogo"
                                    Componentes ("Bloquear")

                                        Exit Do
                                    
                                    End If
                                    
                                Else
                                ' Caso "ContadorTimerMovimento + 1" seja maior que 10,
                                'não há mais recordes além destes; além disso, a pontuação
                                'obtida por este jogador já existe entre os Recordes, sendo
                                'inclusive a última disponível. Portanto, simplesmente
                                'ignora-se a pontuação obtida por este jogador.
                                
                                    Exit Do
                                
                                End If
                            
                            Else
                            'Não sendo igual, pergunta-se se é menor
                            
                                If RecordesJogo(ContadorTimerMovimento, 2) < CLng(lblPontos.Caption) Then
                                ' Sendo menor, considera-se esta como a posição
                                'de recorde do Jogador
                                    
                                    PosicaoRecorde = ContadorTimerMovimento
                                        
                                    fraNovoRecorde.Visible = True
                                    txtNomeJogador.SetFocus
                                    
                                    'Bloqueia os componentes de "frmJogo"
                                    Componentes ("Bloquear")
                                    
                                    Exit Do
                                    
                                End If
                            
                            End If

Continuar_DO:

                            ContadorTimerMovimento = ContadorTimerMovimento + 1
                        
                        Loop
                        
                        Exit Sub
                        
                    End If
                
                    ContadorTimerMovimento = ContadorTimerMovimento + 1
                
                Loop
   
               
        'FIM DOS CÓDIGOS DE AFERIÇÃO DO FIM DE JOGO===========================
            
            Else
            
                'Move os blocos uma posição para baixo
                MoverBloco 40, IndicesEmJogo(1), IndicesEmJogo(2), _
                IndicesEmJogo(3), IndicesEmJogo(4), TipoBlocoEmJogo, _
                PosicaoDoBloco
            
            End If
            
'USO NA DEPURAÇÃO DO SISTEMA ********************************
            
    'Atualiza a exibição da matriz "Jogo" em "txtDebug"
    ExibirValJogo
            
'************************************************************
    
    End If

    Exit Sub
    
Erro:

    MsgBox "Erro encontrado! Erro N°" & Err.Number & "; Descrição: " & Err.Description

    End

End Sub

Private Sub TimerJogo_Timer()
'Realiza o movimento de descida dos Blocos
    
    MoverBlocoParaBaixo = True

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
    
    'Armazena o texto em Português para a variável "MsgBoxGravarRecorde"
    MsgBoxGravarRecorde = "O caracter '%' não é permitido. Digite um nome válido."
    
    'Indica o rótulo utilizado em "cmdNovoJogo" (no caso a
    'posição "0" da variável "cmdNovoJogoTEXTO"
    cmdNovoJogoSTATUS = 0
    
    'Carrega os dados de Recordes do Jogo
    Recordes "CarregarVariavel"
    
    'Armazena a maior pontuação em "lblRecorde"
    lblRecorde.Caption = RecordesJogo(1, 2)
    
    ' Armazena a constante que define uma posição na movimentação
    'dos blocos em "pctJogo"
    UmaPosicao = 375
    
    'Prepara o jogo para ser iniciado
    PrepararJogo
    
    'O menu de término de jogo é desabilitado
    mnuFinalizar.Enabled = False
    
    'Carrega as configurações do jogo
    Configuracoes ("Carregar")
    
    'Indica a altura (Height) correta do Formulário
    frmJogo.Height = 7725
    
'COMANDOS PARA DEPURAÇÃO DO SISTEMA ************************
            
    'Esconde a parte do formulário com os controles de depuração
    frmJogo.Width = 6120
    
    'Já cria uma primeira imagem da variável "Jogo"
    ExibirValJogo
    
'***********************************************************

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Verifica se se deve perguntar acerca do salvamento das configurações
    
    'Impede, a princípio, o descarregamento de "frmJogo" da memória
    Cancel = 1
    
    'Pára os Timers do Jogo
    TimerJogo.Enabled = False
    TimerMovimento.Enabled = False
    
    'Para qualquer possível Música
    TocarMusica Parar
    
    'Verifica se se deve exibir "frmSalvarConfig"
    
    ' Analisa-se o valor do parâmetro na posição número 1 de
    '"ConfigJogo" ("SalvarConfig.Show")
    '   - Se "True", exibe o FORM;
    '   - Se "False", não exibe o FORM.
           
    If ConfigJogo(1, 2) = "True" Then
    'Sendo "True", exibe "frmSalvarConfig"
           
        frmSalvarConfig.Show (vbModal)
           
    Else
    'Sendo "False", não exibe "frmSalvarConfig"; todavia, verifica-se
    'se o usuário já houvera selecionado a opção em "frmSalvarConfig"
    'para que as configurações sejam salvas.
    
        ' Analisa-se o valor do parâmetro na posição número 2 de
        '"ConfigJogo" ("SalvarConfig")
        '   - Se "True", salvam-se as configuções;
        '   - Se "False", não se salvam as configuções.
               
        If ConfigJogo(2, 2) = "True" Then
        'Sendo "True", salvam-se as configuções
               
            Configuracoes ("Salvar")
               
        End If
    
        'Finaliza-se do Jogo
        End
    
    End If
           
End Sub

'===========================================================
'COMANDOS REFERENTES AOS MENUS                             =
'===========================================================
Private Sub mnuNovoJogo_Click()
'Limpa a tela de Jogo e, depois, inicia um
'novo Jogo, habilitando o Timer ("TimerJogo")
       
    'Adiciona o texto "Iniciar" a "cmdNovoJogo"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
       
    'Inicia o Jogo
    IniciarJogo
    
End Sub

Private Sub mnuFinalizar_Click()
'Finaliza o Jogo que estiver em andamento

    TimerJogo.Enabled = False
    TimerMovimento.Enabled = False
    
    'Prepara o jogo para ser novamente inicializado
    PrepararJogo
    
    'O menu de término de jogo é desabilitado
    mnuFinalizar.Enabled = False
                             
    'Habilita os menus selecionadores de Nível de Jogo
    MenuNivel ("Habilitar")

End Sub

Private Sub mnuSair_Click()
'Sai do jogo
    
    Unload Me

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
    
    'Armazena o Estilo do Bloco na variável correspondente
    EstiloBlocos = Index
    
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
    
    'Armazena o Nível do Jogo na variável correspondente
    NivelJogo = Index
    
    'Tira o "Checked" dos menus de Níveis
    Uncheck ("Nível")
    
    SelecionarNivel NivelJogo, False

End Sub

Private Sub mnuIdioma_Click(Index As Integer)
'Troca o idioma do Jogo, utilizando como base a variável
'"Index", através da função "TrocarIdioma()"

    'Tira o "Checked" dos menus de Idioma
    Uncheck ("Idioma")

    Select Case Index
    
        Case 0
            mnuIdioma(0).Checked = True
            
            'Armazena o Idioma do Jogo na variávle correspondente
            IdiomaJogo = "Ptb"
        
        Case 1
            mnuIdioma(1).Checked = True
            
            'Armazena o Idioma do Jogo na variávle correspondente
            IdiomaJogo = "Eng"
            
    End Select
    
    TrocarIdioma (IdiomaJogo)
    
    'Coloca o novo Rótulo em "cmdNovoJogo"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(cmdNovoJogoSTATUS)
    
End Sub

Private Sub mnuMusica_Click()
' Esta SubRotina habilita ou não as músicas do Jogo
'Indica "True" no método Checked do menu selecionado (caso
'seja selecionada a opção) ou "False (caso a opção não seja
'selecionada)

    If mnuMusica.Checked = True Then
    'Pára qualquer música que possa estar sendo executada.
    'além de retirar o "checked" do menu
    
        'Para qualquer possível Música
        TocarMusica Parar
    
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

Private Sub mnuRecordes_Click()
'Exibe "fraRecordes"

    Recordes "CarregarVariavel"
    Recordes "Exibir"

    fraRecordes.Visible = True
    
    'Bloqueia os componentes de "frmJogo"
    Componentes ("Bloquear")

End Sub

'===========================================================

Function IniciarJogo()

'Inicia um novo Jogo, habilitando o Timer ("TimerJogo")
    
    If cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0) Then
        'Inicia o Jogo
        
        Dim TipoBlocoNovo As Integer, EstiloBlocoNovo As Integer
        
        'Prepara o jogo para ser iniciado
        PrepararJogo
        
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
        
        If mnuMusica.Checked = True And MusicaExecutando = True Then

            TocarMusica Pausar

        End If
        
    Else
        'Reinicia o Jogo, após Pausar
        
        TimerJogo.Enabled = True
        TimerMovimento.Enabled = True
        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(1)
        cmdNovoJogoSTATUS = 1
        
        If mnuMusica.Checked = True And MusicaExecutando = True Then

            TocarMusica Resumir

        End If
        
    End If
  
    pctJogo.SetFocus
    
End Function

Function PrepararJogo()
'Prepara o Jogo para ser iniciado

    Dim ContadorNovoJogo1, ContadorNovoJogo2 As Integer
        
    'Limpa as variáveis
    TipoBlocoProx = 0
    EstiloBlocoProx = 0
    
    'Atribue os valores iniciais aos rótulos
    lblPontos.Caption = "0"
    cmdNovoJogo.Caption = cmdNovoJogoTEXTO(0)
    
    'O menu de término de jogo é habilitado
    mnuFinalizar.Enabled = True
        
    'Armazena "False" para todas as posições dos blocos de "pctJogo"
    ContadorNovoJogo1 = 0
    
    Do While ContadorNovoJogo1 <= 179
    
        BlocoEmJogo(ContadorNovoJogo1) = False
    
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop
    
    'Armazena "999" para todas as posições da matriz "EstiloBlocoEmJogo"
    ContadorNovoJogo1 = 0
    
    Do While ContadorNovoJogo1 <= 179
    
        EstiloBlocoEmJogo(ContadorNovoJogo1) = 999
    
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop
        
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
           
    'Volta os blocos de "pctJogo" às posições iniciais e sem figuras anexas
    ContadorNovoJogo1 = 0
    
    Do While ContadorNovoJogo1 <= 179
    
        With imgBloco(ContadorNovoJogo1)
            .Top = 0
            .Left = 0
            .Visible = False
            .Picture = Nothing
        End With
    
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop
    
    'Volta os blocos de "pctProx" às posições iniciais e sem figuras anexas
    ContadorNovoJogo1 = 0
    
    Do While ContadorNovoJogo1 <= 3
    
        With imgBlocoProx(ContadorNovoJogo1)
            .Top = 0
            .Left = 0
            .Visible = False
            .Picture = Nothing
        End With
    
        ContadorNovoJogo1 = ContadorNovoJogo1 + 1
    
    Loop

    'Verifica qual o Nível selecionado para inicialização
    ContadorNovoJogo2 = 0
    
    Do While ContadorNovoJogo2 <= 9
    
        If mnuNivel(ContadorNovoJogo2).Checked = True Then
        
            SelecionarNivel ContadorNovoJogo2, False
        
            Exit Do
        
        End If
   
        ContadorNovoJogo2 = ContadorNovoJogo2 + 1
    
    Loop

End Function

Function Componentes(Acao As String)
'Bloqueia ("Acao = Bloquear") ou desbloqueia ("Acao = Desbloquear")
'alguns componentes de "frmJogo" para impedir que estes
'sejam clicados quando da exibição de "fraRecordes" ou "fraNovoRecorde"

    Select Case Acao
    
        Case "Bloquear"
            
            mnuJogo.Enabled = False
            mnuOpcoes.Enabled = False
            mnuAjuda.Enabled = False
            cmdNovoJogo.Enabled = False
        
        Case "Desbloquear"
        
            mnuJogo.Enabled = True
            mnuOpcoes.Enabled = True
            mnuAjuda.Enabled = True
            cmdNovoJogo.Enabled = True
        
    End Select

End Function

'USO NA DEPURAÇÃO DO SISTEMA ********************************

Private Sub cmdDebugEsconder_Click()
'Esconde a parte do formulário com os controles de depuração

    frmJogo.Width = 6120
    frmJogo.Left = frmJogo.Left + 1860

End Sub

Private Sub cmdDebugScore_Click()
' Acresce "-%0" em todas as linhas do arquivo "score.lst",
'presente na pasta "Data" da pasta de instalação do Jogo

    Dim vLinhaSCORE As Long
    Dim ContadorDebugScore As Integer
    
    vLinhaSCORE = FreeFile
    ContadorDebugScore = 1

    Open App.Path & "\Data\score.lst" For Output As #vLinhaSCORE
            
    'Realiza o loop DEZ vezes
    Do While ContadorDebugScore <= 10
                           
        Print #vLinhaSCORE, "-%0"
                       
        ContadorDebugScore = ContadorDebugScore + 1
                
    Loop
                
    'Fecha o Arquivo Texto
    Close vLinhaSCORE
    
    lblRecorde.Caption = "0"
    
    MsgBox "Recordes iniciais reconfigurados.", vbOKOnly + vbInformation, "Down Blocks DEBUGGER"

End Sub

Private Sub cmdConfigINI_Click()
'Regrava o arquivo "config.ini"

    Dim LinhaArquivoTexto As Long
    
    LinhaArquivoTexto = FreeFile

    'Abre "config.ini" para iserção de dados
    Open App.Path & "\Data\config.ini" For Output As #LinhaArquivoTexto
        
    Print #LinhaArquivoTexto, "SalvarConfig.Show=True"
    Print #LinhaArquivoTexto, "SalvarConfig=False"
    Print #LinhaArquivoTexto, "Musica=True"
    Print #LinhaArquivoTexto, "Sons=True"
    Print #LinhaArquivoTexto, "EstiloBlocos=0"
    Print #LinhaArquivoTexto, "Nivel=0"
    Print #LinhaArquivoTexto, "Idioma=Ptb"
                  
    'Fecha o Arquivo Texto
    Close LinhaArquivoTexto
    
    'Carrega novamente as configurações padrão
    Configuracoes ("Carregar")
    
    MsgBox "Configuração Padrão reestabelecida.", vbOKOnly + vbInformation, "Down Blocks DEBUGGER"

End Sub

'************************************************************
