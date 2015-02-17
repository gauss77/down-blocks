VERSION 5.00
Begin VB.Form frmSalvarConfig 
   BorderStyle     =   0  'None
   Caption         =   "Down Blocks"
   ClientHeight    =   1455
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   3510
   Icon            =   "frmSalvarConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNao 
      Caption         =   "Não"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   690
      Width           =   1095
   End
   Begin VB.CommandButton cmdSim 
      Caption         =   "Sim"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   690
      Width           =   1095
   End
   Begin VB.CheckBox chkNao 
      Caption         =   "Não perguntar novamente"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1130
      Width           =   3255
   End
   Begin VB.Line LineSalvarConfig 
      Index           =   3
      X1              =   3480
      X2              =   3480
      Y1              =   0
      Y2              =   1440
   End
   Begin VB.Line LineSalvarConfig 
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1440
   End
   Begin VB.Line LineSalvarConfig 
      Index           =   1
      X1              =   0
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line LineSalvarConfig 
      Index           =   0
      X1              =   0
      X2              =   3480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image imgPergunta 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "frmSalvarConfig.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "Deseja salvar suas configurações?"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSalvarConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNao_Click()
'Não desejando salvar-se as configurações, nada ocorre
        
    ConfigJogo(2, 2) = "False"
        
    Unload Me

End Sub

Private Sub cmdSim_Click()
'Salvam-se as configurações

    ConfigJogo(2, 2) = "True"
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Verifica se "chkNao" está selecionado
    'Obs.: "chkNao" indica se este FORM deve ser exibido
    'novamente ou não:
    '   - "chkNao.Value = 0" (não selecionado), apenas sai;
    '   - "chkNao.Value = 1" (selecionado), sai do jogo, mas
    '    grava o parâmetro "SalvarConfig.Show" de "config.ini"
    '    como "False"

    If chkNao.Value = 1 Then
    'Impede que este FORM seja exibido novamente
        
       ConfigJogo(1, 2) = "False"
       
    End If
    
    If ConfigJogo(2, 2) = "True" Then
    ' Se foi selecionada a opção "Sim" - Salvar -, salva-se as
    'configurações
    
        Configuracoes ("Salvar")
    
    Else
    ' Se foi selecionada a opção "Não" - Não Salvar - , apenas
    'grava-se o que está em "ConfigJogo" sem aferir as alterações
    
        Configuracoes ("SalvarSemAlterar")
    
    End If
    
    'Finaliza o jogo
    End

End Sub
