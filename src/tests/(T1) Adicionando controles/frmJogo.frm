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
   Begin VB.CommandButton Command1 
      Caption         =   "Remover"
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
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
   Begin VB.Timer Timer1 
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
      Begin VB.Menu mnuEstilo 
         Caption         =   "Estilo do Bloco"
         Begin VB.Menu mnuEstClassico 
            Caption         =   "Clássico"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuEstNovo 
            Caption         =   "Novo"
         End
      End
      Begin VB.Menu mnuEspaco4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIdiomas 
         Caption         =   "Idioma"
         Begin VB.Menu mnuIdiomaPort 
            Caption         =   "Português"
         End
         Begin VB.Menu mnuIdiomaIngles 
            Caption         =   "English"
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
Dim NumObj As Integer
Dim Img As Boolean
Dim NumPct As Integer
Dim PosX As Integer
Dim posY As Integer

Private Sub cmdNovoJogo_Click()

itm = 0

Do While itm < 180

    'Insere as figuras em "ImageListBlocos"
    'Como são sete figuras, o DO será executado 7 vezes
    
    If NumObj = 0 Then
    
        PosX = 0
        posY = 0
    
    End If
    
    NumObj = NumObj + 1
    NumPct = NumPct + 1
    
    If NumPct > 7 Then NumPct = 1
    
    If Img = False Then
    
    Img = True
    
    End If

    Dim pctBlocos As VB.PictureBox
    Set pctBlocos = Controls.Add("VB.PictureBox", "pctBloco" & NumObj, pctJogo)
    With pctBlocos
        .Visible = True
        .Height = 375
        .Width = 375
        .Top = posY
        .Left = PosX
        .Appearance = 0
        .BorderStyle = 0
        .Picture = ImageListBlocos.ListImages(NumPct).Picture
    End With

    PosX = PosX + 375
    
    If PosX = 3750 Then
    
        posY = posY + 375
        PosX = 0
    
    End If

itm = itm + 1


Loop

PosX = 0
posY = 0


End Sub

Private Sub Command1_Click()

Contador = 180

Do While Contador > 0

frmJogo.Controls.Remove "pctbloco" & Contador

Contador = Contador - 1

Loop
End Sub

Private Sub Form_Load()
CarregarImageList (App.Path & "\Blocos\Clássico\")
End Sub

Private Sub mnuEstClassico_Click()
    mnuEstClassico.Checked = True
    mnuEstNovo.Checked = False
    EsvaziarImageList
    CarregarImageList (App.Path & "\Blocos\Clássico\")
  
End Sub

Private Sub mnuEstNovo_Click()
    mnuEstClassico.Checked = False
    mnuEstNovo.Checked = True
    EsvaziarImageList
    CarregarImageList (App.Path & "\Blocos\Novo\")
End Sub

Private Sub mnuIdiomaIngles_Click()

    TrocarIdioma ("Eng")
    
End Sub

Private Sub mnuIdiomaPort_Click()

    TrocarIdioma ("Ptb")

End Sub

Private Sub mnuMusica_Click()

    If mnuMusica.Checked = True Then

        mnuMusica.Checked = False
        
    Else
    
        mnuMusica.Checked = True
    
    End If

End Sub

'COMANDOS REFERENTES AOS MENUS =============================

Private Sub mnuSair_Click()

    'Finaliza o jogo
    End

End Sub


Private Sub mnuSobre_Click()

    frmSobre.Show (vbModal)

End Sub

'===========================================================
Private Sub mnuSons_Click()

End Sub

Function CarregarImageList(Endereco As String)

    Dim Contador As Integer

    Contador = 1

    Do While Contador <= 7
    
        ImageListBlocos.ListImages.Add Contador, "Bloco" & Contador, LoadPicture(Endereco & Contador & ".jpg")
    
        Contador = Contador + 1
    
    Loop

End Function

Function EsvaziarImageList()

    Dim Contador As Integer

    Contador = 7

    Do While Contador <> 0
    
        ImageListBlocos.ListImages.Remove (Contador)
    
        Contador = Contador - 1
    
    Loop

End Function


Function TrocarIdioma(Idioma As String)

    Dim vControle As String, vLabel As String

    vLinhaTXT = FreeFile
    
    Open App.Path & "\Idioma\" & Idioma & ".lng" For Input As vLinhaTXT 'Abre o arquivo texto
    
    'Realiza o loop enquanto não for fim do arquivo
    Do While Not EOF(vLinhaTXT)
        
        'Lê a linha do arquivo texto onde o cursor está
        Line Input #vLinhaTXT, linha
   
        vControle = Left(linha, InStr(1, linha, "=") - 1)
        vLabel = Right(linha, Len(linha) - InStr(1, linha, "="))
        
        Select Case vControle
        
            Case "mnuJogo"
                mnuJogo.Caption = vLabel
                
            Case "mnuNovoJogo"
                mnuNovoJogo.Caption = vLabel
            
            Case "mnuSair"
                mnuSair.Caption = vLabel
            
            Case "mnuOpcoes"
                mnuOpcoes.Caption = vLabel
            
            Case "mnuMusica"
                mnuMusica.Caption = vLabel
            
            Case "mnuSons"
                mnuSons.Caption = vLabel
            
            Case "mnuEstilo"
                mnuEstilo.Caption = vLabel
                
            Case "mnuEstClassico"
                mnuEstClassico.Caption = vLabel
            
            Case "mnuEstNovo"
                mnuEstNovo.Caption = vLabel
            
            Case "mnuIdiomas"
                mnuIdiomas.Caption = vLabel
            
            Case "mnuAjuda"
                mnuAjuda.Caption = vLabel
            
            Case "mnuConteudo"
                mnuConteudo.Caption = vLabel
            
            Case "mnuSobre"
                mnuSobre.Caption = vLabel
            
            Case "lblTXTProx"
                lblTXTProx.Caption = vLabel
                
            Case "lblTXTPontos"
                lblTXTPontos.Caption = vLabel
            
            Case "lblTXTNivel"
                lblTXTNivel.Caption = vLabel
            
            Case "cmdNovoJogo"
                cmdNovoJogo.Caption = vLabel
            
            Case "frmSobre"
                frmSobre.Caption = vLabel
            
            Case "lblTeste"
                frmSobre.lblTeste.Caption = vLabel

        End Select

    Loop
        
    'Fecha o Arquivo Texto
    Close vLinhaTXT


End Function
