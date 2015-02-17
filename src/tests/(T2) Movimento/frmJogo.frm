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
      Height          =   855
      Left            =   4080
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   600
      Width           =   1815
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
    
        TimerJogo.Enabled = False
        TimerMovimento.Enabled = False
        cmdNovoJogo.Caption = cmdNovoJogoTEXTO(2)
        cmdNovoJogoSTATUS = 2
        
    Else
    
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
