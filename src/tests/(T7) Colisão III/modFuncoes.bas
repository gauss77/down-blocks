Attribute VB_Name = "modFuncoes"
'==============================================================
'Este m�dulo cont�m as fun��es utilizadas no Jogo "Down Blocks"
'2005 - Desenvolvido por Andr� Martins
'Contato: andrel_martins@yahoo.com.br
'==============================================================

Global cmdNovoJogoTEXTO(3) As String ' Em suas tr�s posi��es s�o
                                     'armazenados os textos refe-
                                     'rentes aos r�tulos do con-
                                     'trole "cmdNovoJogo", sendo:
                                     ' - 0: Iniciar
                                     ' - 1: Pausar
                                     ' - 2: Continuar

Global BlocoEmJogo(179) As Boolean ' Indica se os blocos est�o
                                   'em uso em "pctJogo" ou n�o

Global Jogo(18, 10) As Integer ' Indica as posi��es dispon�veis
                               'em "pctJogo", sendo a representa��o
                               'matem�tica deste container
                               ' O valor "999" indica que n�o h�
                               'nenhum bloco nesta posi��o de
                               '"pctjogo". Qualquer valor diferente
                               'disso (que ir� variar de "0" a "179")
                               'indica o bloco (o �ndice de uma das
                               'PictureBox existentes naquele container)
                               'que est� nesta posi��o
                    
Global UmaPosicao As Integer ' Armazena o valor "375", que
                             'indica o valor m�nimo para a
                             'movimenta��o de um bloco em
                             '"pctJogo", para qualquer dire��o
                                 
Global IndicesEmJogo(4) As Integer ' Armazena o �ndice dos blocos
                                   'que estiverem em primeiro plano
                                   'no jogo (blocos que est�o aptos
                                   'a serem movidos em "pctJogo")
                                   ' O que define a quantidade de
                                   '�ndices armazenados � o tipo
                                   'do bloco, representado pela
                                   'vari�vel "TipoBlocoEmJogo"
                                   
Type Posicao             ' Tipo de vari�vel que possui duas
    PosicaoX As Integer  'possiblidadesde inser��o de dados
    PosicaoY As Integer  '("PosicaoX" e "PosicaoY"), utilizado
End Type                 'para criar vari�veis que armazenaram
                         'as coordenadas cartesianas X e Y do
                         '�ndice de um bloco na matriz "Jogo"

Type PosicoesParaIndices ' Tipo de vari�vel que armazena �ndices
    idx1 As Integer      'das 4 posi��es dos blocos presentes
    idx2 As Integer      'Jogo. � utitilizada em diversas fun��es
    idx3 As Integer
    idx4 As Integer
End Type

Public Declare Function TocarSom Lib "WINMM.DLL" Alias _
 "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Function CarregarImageList(Endereco As String)
'============================================================
' Esta fun��o carrega as imagens utilizadas nos blocos para
'um ImageList e, caso j� haja dados no ImageList, deleta-os
'antes de inserir as novas imagens
'============================================================
    
    Dim Contador As Integer

    ' Se o contador do ImageList for maios que "0", indica que
    'h� imagens armazenadas. Assim, exclue-as antes de iniciar
    'o carregamento das novas imagens
    If frmJogo.ImageListBlocos.ListImages.Count > 0 Then

        Contador = 7
    
        Do While Contador > 0
        
            frmJogo.ImageListBlocos.ListImages.Remove (Contador)
        
            Contador = Contador - 1
        
        Loop
    
    End If

    'Insere as novas imagens no contador, segundo a sele��o do
    'usu�rio, indica pela vari�vel "Endereco"
    Contador = 1

    Do While Contador <= 7
    
        frmJogo.ImageListBlocos.ListImages.Add Contador, "Bloco" & Contador, LoadPicture(Endereco & Contador & ".jpg")
    
        Contador = Contador + 1
    
    Loop

End Function

Function TrocarIdioma(Idioma As String)
'============================================================
' Esta fun��o troca o idioma dos menus e labels dos formul�-
'rios que comp�em o Jogo
'============================================================

    Dim vControle As String, vLabel As String
    Dim Menu As Object
    Dim Formulario As String

    vLinhaTXT = FreeFile
    
    ' Abre o arquivo de Idiomas (.lng), que est� na pasta "\Idioma"
    'segundo a sele��o do usu�rio (vari�vel "Idioma"
    Open App.Path & "\Idioma\" & Idioma & ".lng" For Input As vLinhaTXT 'Abre o arquivo texto
    
    'Realiza o loop enquanto n�o for fim do arquivo
    Do While Not EOF(vLinhaTXT)
        
        'L� a linha do arquivo texto onde o cursor est�
        Line Input #vLinhaTXT, Linha
   
        ' Verifica se o texto da corrente linha do arquivo �
        'indica��o de nome de formul�rio
        If Left(Linha, 1) = "*" Then
        
            Formulario = Linha
            ' L� a pr�xima linha do arquivo texto (que indica
            'um controle do corrente formul�rio)
            Line Input #vLinhaTXT, Linha
            
        End If
   
        'Verifica de qual formul�rio pertence o atual componente
        Select Case Formulario
        
            Case "*Variaveis*"
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))
                Select Case Left(vControle, Len(vControle) - 1)
                
                    Case "cmdNovoJogoTEXTO"
                        cmdNovoJogoTEXTO(CInt(Right(vControle, 1))) = vLabel
        
                End Select
        
            Case "*frmJogo*"
               
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o r�tulo do controle
                '*Usa como base para separa��o da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))
                
                ' Verifica se o �ltimo caractere do nome do controle
                '(valor de "vControle") � um n�mero
                If IsNumeric(Right(vControle, 1)) = True Then
        
                    ' Sendo n�mero, modifica o r�tulo do corrente objeto
                    'com o idioma selecionado mas, ao inv�s de usar o
                    'valor de "vControle" como base para designar o
                    'objeto, retira o n�mero da vari�vel e usa este
                    'como substrato para a forma��o do �ndice do corrente
                    'objeto
                    Set Menu = frmJogo(Left(vControle, Len(vControle) - 1))
                    Menu(CInt(Right(vControle, 1))).Caption = vLabel
        
                Else
                
                    'Modifica o r�tulo do corrente objeto o idioma selecionado
                    Set Menu = frmJogo(vControle)
                    Menu.Caption = vLabel
        
                End If
         
            Case "*frmSobre*"
    
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o r�tulo do controle
                '*Usa como base para separa��o da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))

                If IsNumeric(Right(vControle, 1)) = True Then
        
                    ' Sendo n�mero, modifica o r�tulo do corrente objeto
                    'com o idioma selecionado mas, ao inv�s de usar o
                    'valor de "vControle" como base para designar o
                    'objeto, retira o n�mero da vari�vel e usa este
                    'como substrato para a forma��o do �ndice do corrente
                    'objeto
                    Set Menu = frmJogo(Left(vControle, Len(vControle) - 1))
                    Menu(CInt(Right(vControle, 1))).Caption = vLabel
        
                Else
                
                    'Verifica se o controle atual � o pr�prio Formul�rio
                    If vControle = "frmSobre" Then
                
                        'Modifica o r�tulo do Formul�rio
                        frmSobre.Caption = vLabel
                        
                    Else
                    
                        'Modifica o r�tulo do corrente objeto o idioma selecionado
                        Set Menu = frmSobre(vControle)
                        Menu.Caption = vLabel
                        
                    End If
        
                End If
                       
        End Select
    
    Loop
        
    'Fecha o Arquivo Texto
    Close vLinhaTXT

End Function

Function Uncheck(Menu As String)
'=============================================================
' Esta fun��o retira os "Checks" existentes na �rea do menu
'que cont�m os Estilos dos Blocos, os Idiomas ou os N�veis do
'Jogo (segundo sele��o do usu�rio indicada atrav�s da vari�vel
'"Menu"
'=============================================================

    Dim Contador As Integer

    Select Case Menu
    
        Case "Estilo"
            Contador = 0
        
            'Realiza-se um la�o que dever� ser executado at� que se
            'atinja o n�mero de Estilos de Blocos existentos no menu
            Do While Contador < 2
        
                'Indica "False" para o m�todo "Checked" do controle
                frmJogo.mnuEst(Contador).Checked = False
        
                'Adiciona 1 � vari�vel contadora
                Contador = Contador + 1
        
            Loop
            
        Case "Idioma"
            Contador = 0
        
            'Realiza-se um la�o que dever� ser executado at� que se
            'atinja o n�mero de Estilos de Blocos existentos no menu
            Do While Contador < 2
        
                'Indica "False" para o m�todo "Checked" do controle
                frmJogo.mnuIdioma(Contador).Checked = False
        
                'Adiciona 1 � vari�vel contadora
                Contador = Contador + 1
                
            Loop
                
        Case "N�vel"
            Contador = 0
        
            'Realiza-se um la�o que dever� ser executado at� que se
            'atinja o n�mero de Estilos de Blocos existentos no menu
            Do While Contador < 10
        
                'Indica "False" para o m�todo "Checked" do controle
                frmJogo.mnuNivel(Contador).Checked = False
        
                'Adiciona 1 � vari�vel contadora
                Contador = Contador + 1
                
            Loop
        
    End Select

End Function

Function GerarBlocos(Container As String, TipoBloco As Integer, EstiloBloco As Integer)
'=============================================================
' Esta fun��o gera os blocos aleat�rios que ser�o utilizados
'durante o Jogo
' H� 5 tipos de Blocos que podem ser gerados, com 7 cores
'distintas para a forma��o destes (asteriscos representam os
'blocos):

'   - ****; representado pelo n�mero "0"

'   - ***; representado pelo n�mero "1"

'      *
'   - ***; representado pelo n�mero "2"

'       *
'   - ***; representado pelo n�mero "3"

'      **
'   - ** ; representado pelo n�mero "4"

'     **
'   - **; representado pelo n�mero "5"
'
' As vari�veis da fun��o indicam:
'   - Container: indica a PictureBox que conter� os Blocos
'             ("pctJogo" ou "pctProx");
'   - TipoBloco: indica o Tipo do Bloco (� um n�mero aleat�rio
'             gerado atrav�s de "Random(4)"
'   - EstiloBloco: indica o Estilo do Bloco (uma imagem sele-
'             cionada aleatoriamente atrav�s de "Random(7)"
'             em ImageListBlocos
'
'=============================================================

    Dim Contador As Integer
    Dim PosX As Integer, PosY As Integer
    Dim pctObj As Object
    Dim IndexBloco As Integer 'Armazena o n�mero do �ndice
    
    'Indica qual PictureBox utilizar, atrav�s do Container
    If Container = "pctJogo" Then
    
        'Armazena "999" para as posi��es dos �ndices dos blocos no Jogo
        IndicesEmJogo(1) = 999
        IndicesEmJogo(2) = 999
        IndicesEmJogo(3) = 999
        IndicesEmJogo(4) = 999
    
        'Indica a qual PictureBox se referenciar� "pctObj"
        Set pctObj = frmJogo.imgBloco
    
    Else
        
        'Indica a qual PictureBox se referenciar� "pctObj"
        Set pctObj = frmJogo.imgBlocoProx
    
        ' Esconde todos os Blocos utilizados em "pctProx" e
        'retorna-os a suas posi��es iniciais
        Contador = 0
        
        Do While Contador < 4
        
            With pctObj(Contador)
                .Visible = False
                .Top = 0
                .Left = 0
            End With
        
            Contador = Contador + 1
        
        Loop
    
    End If

    ' Assegura que o valor de EstiloBloco esteja sempre entre
    '"1" e "7"
    If EstiloBloco = 0 Then EstiloBloco = 1

    Select Case TipoBloco
    
        Case 0
            'Monta um conjunto de Blocos do tipo ****, no Container escolhido
                
            Contador = 1
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es iniciais do primeiro bloco
                PosX = 3 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posi��es iniciais do primeiro bloco
                PosX = 150
                PosY = 237
            
            End If

            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 0
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If

                Contador = Contador + 1
                
            Loop
                
        Case 1
            'Monta um conjunto de Blocos do tipo ***, no Container escolhido
        
            Contador = 1
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es iniciais do primeiro bloco
                PosX = 3 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posi��es iniciais do primeiro bloco
                PosX = 337
                PosY = 237
            
            End If

            Do While Contador <= 3
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 1

                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If
                
                Contador = Contador + 1
                
            Loop
        
        Case 2
            '                                     *
            'Monta um conjunto de Blocos do tipo ***, no Container escolhido
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es iniciais do primeiro bloco
                PosX = 4 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posi��es iniciais do primeiro bloco
                PosX = 712
                PosY = 50
            
            End If
            
            IndexBloco = ProxIndexBloco(Container)
            
            'Insere o primeiro bloco
            pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
            pctObj(IndexBloco).Top = PosY
            pctObj(IndexBloco).Left = PosX
            pctObj(IndexBloco).Visible = True

            If Container = "pctJogo" Then

                InserirNaMatriz IndexBloco, 1, 2
                
                ' Armazena o �ndice deste bloco para que este
                'seja movimentado em "pctJogo"
                IndicesEmJogo(1) = IndexBloco
                    
            End If
            
            'Prepara-se para inserir os pr�ximos tr�s blocos
            'Definem-se as posi��es iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es
                PosX = 3 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posi��es
                PosX = 337
                PosY = 425
                
            End If
            
            ' Contador inicia-se em 2 por j� se haver inserido
            'o primeiro bloco
            Contador = 2

            'Inserem-se os pr�ximos 3 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 2
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If
                
                Contador = Contador + 1
                
            Loop
        
        Case 3
            '                                      *
            'Monta um conjunto de Blocos do tipo ***, no Container escolhido
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es iniciais do primeiro bloco
                PosX = 5 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posi��es iniciais do primeiro bloco
                PosX = 1087
                PosY = 50
            
            End If
            
            IndexBloco = ProxIndexBloco(Container)
            
            'Insere o primeiro bloco
            pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
            pctObj(IndexBloco).Top = PosY
            pctObj(IndexBloco).Left = PosX
            pctObj(IndexBloco).Visible = True

            If Container = "pctJogo" Then

                InserirNaMatriz IndexBloco, 1, 3
                
                ' Armazena o �ndice deste bloco para que este
                'seja movimentado em "pctJogo"
                IndicesEmJogo(1) = IndexBloco
                    
            End If
            
            'Prepara-se para inserir os pr�ximos tr�s blocos
            'Definem-se as posi��es iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es
                PosX = 3 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posi��es
                PosX = 337
                PosY = 425
                
            End If
            
            ' Contador inicia-se em 2 por j� se haver inserido
            'o primeiro bloco
            Contador = 2

            'Inserem-se os pr�ximos 3 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 3
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If

                Contador = Contador + 1
                
            Loop
            
        Case 4
            '                                     **
            'Monta um conjunto de Blocos do tipo ** , no Container escolhido
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es iniciais do primeiro bloco
                PosX = 4 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posi��es iniciais do primeiro bloco
                PosX = 715
                PosY = 50
            
            End If
            
            'Inserem-se os 2 primeiros blocos
            Contador = 1
            
            Do While Contador <= 2
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 4
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If

                Contador = Contador + 1
                
            Loop
            
            'Prepara-se para inserir os pr�ximos 2 blocos
            'Definem-se as posi��es iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es
                PosX = 3 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posi��es
                PosX = 337
                PosY = 425
                
            End If

            'Inserem-se os pr�ximos 2 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 4
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If

                Contador = Contador + 1
                
            Loop
            
        Case 5
            '                                    **
            'Monta um conjunto de Blocos do tipo **, no Container escolhido

            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es iniciais do primeiro bloco
                PosX = 4 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posi��es iniciais do primeiro bloco
                PosX = 525
                PosY = 50
            
            End If
            
            'Inserem-se os 2 primeiros blocos
            Contador = 1
            
            Do While Contador <= 2
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 5
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If
                
                Contador = Contador + 1
                
            Loop
            
            'Prepara-se para inserir os pr�ximos 2 blocos
            'Definem-se as posi��es iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posi��es
                PosX = 4 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posi��es
                PosX = 525
                PosY = 425
                
            End If

            'Inserem-se os pr�ximos 2 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 5
                    
                    ' Armazena o �ndice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If
                
                Contador = Contador + 1
                
            Loop

    End Select

'USO NA DEPURA��O DO SISTEMA ********************************
    
    'Atualiza a exibi��o da matriz "Jogo" em "txtDebug"
    ExibirValJogo
    
'************************************************************
    
End Function

Function Random(Base_de_Calculo As Integer, TipoDeAleatoriedade As Integer)
'=======================================================================
'  Esta Fun��o, com base nos segundos do rel�gio interno do computador,
' seleciona uma das tr�s possibilidades de c�lculo de aleatoriedade e
' realiza a sele��o de um n�mero aleat�rio
'=======================================================================
    
    Dim vRnd As Integer
    
    Select Case TipoDeAleatoriedade
    
        Case Is < 20
            vRnd = CInt((((((Second(Time) + Hour(Time) + Day(Date) + Month(Date) + Right(Second(Time) * Hour(Time), 1)) * 314158 + 1) Mod 10000) / 10000) * (Base_de_Calculo + 1)) Mod (Base_de_Calculo + 1))
             
        Case Is > 200
            vRnd = CInt((((((Second(Time) + Hour(Time) + Day(Date) + Month(Date) + Left(Month(Date) * Hour(Time) + Second(Time), 1)) * 314158 + 1) Mod 10000) / 10000) * (Base_de_Calculo + 1)) Mod (Base_de_Calculo + 1))
    
        Case Else
            vRnd = CInt((((((Second(Time) + Hour(Time) + Day(Date) + Month(Date)) * 314158 + 1) Mod 10000) / 10000) * (Base_de_Calculo + 1)) Mod (Base_de_Calculo + 1))
    
    End Select
    
    Random = vRnd

End Function

Function ProxIndexBloco(ContainerDoBloco As String) As Integer
'=======================================================================
'  Esta Fun��o indica qual � o pr�ximo �ndice (que se referenciar� a um
'bloco durante o jogo) a ser utilizado
'  A vari�vel "ContainerDoBloco" indica qual � o container de onde os blocos
'dever�o ser selecionados ("pctJogo" ou "pctProx"), uma vez que os
'blocos destes seguem numera��es diferentes.

'  Quando o container for "pctJogo":
'  A fun��o utiliza os dados da vari�vel "BlocoEmJogo", que armazena
'"True" para blocos que est�o em uso no Jogo (inseridos no container
'"pctJogo") e "False" para blocos que n�o est�o em uso. Atrav�s da
'indica��o da vari�vel, a fun��o busca o pr�ximo bloco dispon�vel
'para utiliza��o a partir do �ndice "0", selecionando como bloco
'dispon�vel aquele que tem como dado armazenado "False"

'  Quando o container for "pctProx":
'  Verificar-se-� a posi��o dos blocos de "pctProx" e aquele que ainda
'estiver em sua posi��o incial "Top = 0" e "Left = 0" ser� utilizado
'=======================================================================
    
    Dim Contador As Integer
    
    Select Case ContainerDoBloco
    
        Case "pctJogo"
    
            Contador = 0
            
            Do While Contador < 180
            
                If BlocoEmJogo(Contador) = False Then
                
                    ProxIndexBloco = Contador
                    
                    'Indica que este bloco est� em uso no Jogo
                    BlocoEmJogo(Contador) = True
                    Exit Do
                
                End If
                
                Contador = Contador + 1
            
            Loop
            
        Case "pctProx"
        
            Contador = 0
            
            Do While Contador < 4

                If frmJogo.imgBlocoProx(Contador).Top = 0 Then
                
                    ProxIndexBloco = Contador
                    Exit Do
                
                End If
                
                Contador = Contador + 1
            
            Loop

    End Select

End Function

Function InserirNaMatriz(Indice As Integer, Posicao As Integer, Tipo_do_Bloco As Integer)
'=========================================================================
'  Esta Fun��o insere as posi��es do bloco com "Indice" e "TipoBloco"
'selecionados na matriz "Jogo" (v�lido quando da inser��o dos blocos
'em "pctJogo"), levando-se em conta a "Posicao" de cada bloco no conjunto.
'=========================================================================

'  Obs.: Os n�meros ap�s as representa��es dos blocos nas descri��es do CASE
'abaixo indicam a "Posicao" de cada bloco no conjunto

    Select Case Tipo_do_Bloco
    
        Case 0
            'Insere um conjunto de Blocos do tipo **** ( 1234 ) na matriz "Jogo"
            'EIXO DO BLOCO: posi��o "2"
            
            Select Case Posicao
                
                Case 1
                    Jogo(1, 4) = Indice

                Case 2
                    Jogo(1, 5) = Indice
                
                Case 3
                    Jogo(1, 6) = Indice
                
                Case 4
                    Jogo(1, 7) = Indice

            End Select
                
        Case 1
            'Insere um conjunto de Blocos do tipo *** ( 123 ) na matriz "Jogo"
            'EIXO DO BLOCO: posi��o "2"
            
            Select Case Posicao
                
                Case 1
                    Jogo(1, 4) = Indice
                
                Case 2
                    Jogo(1, 5) = Indice
                
                Case 3
                    Jogo(1, 6) = Indice
                    
            End Select
        
        Case 2
            '                                      *     1
            'Insere um conjunto de Blocos do tipo *** ( 234 ) na matriz "Jogo"
            'EIXO DO BLOCO: posi��o "3"
            
            Select Case Posicao
                
                Case 1
                    Jogo(1, 5) = Indice
                
                Case 2
                    Jogo(2, 4) = Indice
                
                Case 3
                    Jogo(2, 5) = Indice
                    
                Case 4
                    Jogo(2, 6) = Indice
                    
            End Select
        
        Case 3
            '                                       *     1
            'Insere um conjunto de Blocos do tipo *** ( 234 ) na matriz "Jogo"
            'EIXO DO BLOCO: posi��o "3"
            
            Select Case Posicao
                
                Case 1
                    Jogo(1, 6) = Indice
                
                Case 2
                    Jogo(2, 4) = Indice
                
                Case 3
                    Jogo(2, 5) = Indice
                    
                Case 4
                    Jogo(2, 6) = Indice
                    
            End Select
            
        Case 4
            '                                      **    12
            'Insere um conjunto de Blocos do tipo **  ( 34  ) na matriz "Jogo"
            'EIXO DO BLOCO: posi��o "4"
            
            Select Case Posicao
                
                Case 1
                    Jogo(1, 5) = Indice
                
                Case 2
                    Jogo(1, 6) = Indice
                
                Case 3
                    Jogo(2, 4) = Indice
                    
                Case 4
                    Jogo(2, 5) = Indice
                    
            End Select
            
        Case 5
            '                                     **    12
            'Insere um conjunto de Blocos do tipo **  ( 34 ) na matriz "Jogo"
            'EIXO DO BLOCO: <n�o possui eixo (n�o h� rota��o)>
            
            Select Case Posicao
                
                Case 1
                    Jogo(1, 5) = Indice
                
                Case 2
                    Jogo(1, 6) = Indice
                
                Case 3
                    Jogo(2, 5) = Indice
                    
                Case 4
                    Jogo(2, 6) = Indice
                    
            End Select
            
    End Select
    
End Function

Function LocalizarBlocoNaMatriz(Indice_do_Bloco As Integer) As Posicao
'=========================================================================
'  Esta Fun��o localiza a posi��o (X,Y) de um bloco na matriz, atrav�s da
'indica��o de seu �ndice ("Indice_do_Bloco")
'=========================================================================

    Dim Contador1, Contador2 As Integer
    
    Contador1 = 1 'Indica as linhas da matriz
    Contador2 = 1 'Indica as colunas da matriz
    
    Do While Contador1 <= 18

        Do While Contador2 <= 10
        
            If Jogo(Contador1, Contador2) = Indice_do_Bloco Then
            
                LocalizarBlocoNaMatriz.PosicaoX = Contador2
                LocalizarBlocoNaMatriz.PosicaoY = Contador1
            
                Contador1 = 18
                Exit Do

            End If
        
            Contador2 = Contador2 + 1

        Loop

        Contador2 = 1
        Contador1 = Contador1 + 1

    Loop

End Function

Function VerificarLinhasCompletas() As Integer
'=========================================================================
'  Esta Fun��o verifica cada linha na matriz, observando se todas as
'posi��es de coluna (eixo X) est�o com valores diferentes de "999", o que
'indica uma linha completa, e retorna o n�mero desta linha
' Obs.: inicia-se a verica��o a partir da �ltima linha; retorna-se "999"
'caso nenhuma linha esteja completa
'=========================================================================

    Dim LinhaMatriz As Integer
    Dim ColunaMatriz As Integer
    Dim LinhaCompleta As Boolean
    
    LinhaMatriz = 18
    ColunaMatriz = 1
    
    Do While LinhaMatriz >= 1
    
        LinhaCompleta = True
    
        Do While ColunaMatriz <= 10
        
            If Jogo(LinhaMatriz, ColunaMatriz) = 999 Then
            
                LinhaCompleta = False
                Exit Do
                
            End If
        
            ColunaMatriz = ColunaMatriz + 1
        
        Loop
        
        If LinhaCompleta = True Then
        
            Exit Do
            
        End If
        
        ColunaMatriz = 1
        LinhaMatriz = LinhaMatriz - 1
    
    Loop
    
    If LinhaCompleta = True Then
        
        VerificarLinhasCompletas = LinhaMatriz
            
    Else
        
        VerificarLinhasCompletas = 999
        
    End If
   
End Function

Function ReordenarJogo(Linha_da_Matriz As Integer)
'=========================================================================
'  Esta Fun��o reordena os blocos e os valores a estes relacionados na
'matriz "Jogo" atrav�s de "Linha_da_Matriz", que indica a linha que dever�
'ser exclu�da (e os blocos acima desta movidos para esta posi��o)
'=========================================================================

    Dim Contador_da_Reordem As Integer
    Dim Contador_das_Colunas As Integer

    'Inicia um DO que executar-se-� durante "Linha_da_Matriz - 1" vezes
    Contador_da_Reordem = (Linha_da_Matriz - 1)
    Contador_das_Colunas = 1
    
    Do While Contador_da_Reordem >= 1
    
        Do While Contador_das_Colunas <= 10
        
            'Verifica se h� bloco nesta posi��o
            If Jogo(Contador_da_Reordem, 1) <> 999 Then
            
                'Move o bloco uma posi��o abaixo
                frmJogo.imgBloco(Jogo(Contador_da_Reordem, 1)).Top = Contador_da_Reordem * UmaPosicao
                'Indica a nova posi��o do bloco na matrtiz
                Jogo(Linha_da_Matriz, 1) = Jogo(Contador_da_Reordem, 1)
                'Armazena "999" na antiga posi��o do bloco na matriz
                Jogo(Contador_da_Reordem, 1) = 999
            
            End If
        
            Contador_das_Colunas = Contador_das_Colunas + 1
        
        Loop
    
        Contador_das_Colunas = 1
        Contador_da_Reordem = Contador_da_Reordem - 1
    
    Loop

End Function

Function MoverBloco(Direcao As Integer, I1 As Integer, I2 As Integer, _
 I3 As Integer, I4 As Integer)
'=========================================================================
'  Esta Fun��o move os blocos de �ndices "I1", "I2", "I3" e "I4" para:
'       - Baixo ("Direcao = 40")
'       - Esquerda ("Direcao = 37")
'       - Direita ("Direcao = 39")
'=========================================================================

    Dim Blocos_a_mover(4) As Integer
    Dim Contador_MoverBloco As Integer
    Dim Posicao_Bloco_Matriz As Posicao
    'Armazenam as varia��es de posi��es para a manipula��o na matriz "Jogo"
    Dim PosicaoMatriz_X As Integer, PosicaoMatriz_Y As Integer
    'Armazenam as varia��es de posi��es para a manipula��o dos blocos
    Dim PosicaoBloco_X As Integer, PosicaoBloco_Y As Integer

    Select Case Direcao
    
        Case 40
            'Move os blocos uma posi��o para baixo
     
            PosicaoMatriz_X = 0
            PosicaoMatriz_Y = 1
            PosicaoBloco_X = 0
            PosicaoBloco_Y = UmaPosicao
            
            ' Ao mover-se para a direita, como h� acr�scimo
            'de UMA posi��o, deve-se mover primeiramente o
            '�LTIMO bloco do conjunto, e depois os demais,
            'em ordem decrescente
            Blocos_a_mover(1) = I1
            Blocos_a_mover(2) = I2
            Blocos_a_mover(3) = I3
            Blocos_a_mover(4) = I4
        
        Case 37
            'Move os blocos uma posi��o para a esquerda
        
            PosicaoMatriz_X = -1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = -UmaPosicao
            PosicaoBloco_Y = 0
            
            ' Ao mover-se para a esquerda, como h� diminui��o
            'de UMA posi��o, deve-se mover primeiramente o
            'PRIMEIRO bloco do conjunto, e depois os demais,
            'em ordem crescente
            Blocos_a_mover(1) = I4
            Blocos_a_mover(2) = I3
            Blocos_a_mover(3) = I2
            Blocos_a_mover(4) = I1
        
        Case 39
            'Move os blocos uma posi��o para a direita
            
            PosicaoMatriz_X = 1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = UmaPosicao
            PosicaoBloco_Y = 0
            
            ' Ao mover-se para a baixo, como h� acr�scimo
            'de UMA posi��o, deve-se mover primeiramente o
            '�LTIMO bloco do conjunto, e depois os demais,
            'em ordem decrescente
            Blocos_a_mover(1) = I1
            Blocos_a_mover(2) = I2
            Blocos_a_mover(3) = I3
            Blocos_a_mover(4) = I4
            
    End Select
    
    Contador_MoverBloco = 4
        
    Do While Contador_MoverBloco >= 1
            
        If Blocos_a_mover(Contador_MoverBloco) <> 999 Then
    
            'Move o bloco uma posi��o abaixo na matriz
            Posicao_Bloco_Matriz = LocalizarBlocoNaMatriz(Blocos_a_mover(Contador_MoverBloco))
            Jogo(Posicao_Bloco_Matriz.PosicaoY + PosicaoMatriz_Y, Posicao_Bloco_Matriz.PosicaoX + PosicaoMatriz_X) = Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX)
            Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX) = 999
                    
            'Move o bloco 1 posi��o abaixo em "pctJogo"
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top + PosicaoBloco_Y)
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left + PosicaoBloco_X)
                
        End If
            
        Contador_MoverBloco = Contador_MoverBloco - 1
            
    Loop

End Function

Function ArmazenarPosicoes(Direcao_Movimento As Integer, Tipo_BlocoEmJogo As Integer, _
PosicaoAtual_BlocoEmJogo As Integer) As PosicoesParaIndices
'=========================================================================
'  Esta Fun��o indica quais posi��es de �ndices de um determinado
'"Tipo_BlocoEmJogo" devem ser verificadas, usando, para tal, os dados da
'"Direcao_Movimento" que deve ser analisada e da "PosicaoAtual_BlocoEmJogo"
'=========================================================================

    'Primeiramente, verifica qual a dire��o que deve ser analisada
    Select Case Direcao_Movimento
    
        Case 40 'Descida
        
            ' Verifica qual o tipo de Bloco
            Select Case Tipo_BlocoEmJogo
            
                Case 0
                    'Blocos do tipo ****
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica todas as posi��es abaixo do
                            'conjunto de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 4
                        
                        Case 1
                            ' Verifica apenas a posi��o diretamente
                            'abaixo do bloco "4" do conjunto
                            ArmazenarPosicoes.idx1 = 4
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 1
                    'Blocos do tipo ***
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica todas as posi��es abaixo do
                            'conjunto de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas a posi��o diretamente
                            'abaixo do bloco "3" do conjunto
                            ArmazenarPosicoes.idx1 = 3
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 2
                    '                *
                    'Blocos do tipo ***
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "1", "2" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 3
                    '                 *
                    'Blocos do tipo ***
        
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "1", "2" e "3"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posi��es diretamente
                            'abaixo dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 4
                    '                **
                    'Blocos do tipo **
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica apenas as posi��es diretamente
                            'abaixo dos blocos "3" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 3
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posi��es diretamente
                            'abaixo dos blocos "2" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 5
                    '               **
                    'Blocos do tipo **
                    
                    ' Neste tipo de bloco, apenas as posi��es "3" e "4"
                    's�o sempre verificadas
                    ArmazenarPosicoes.idx1 = 3
                    ArmazenarPosicoes.idx2 = 4
                    ArmazenarPosicoes.idx3 = 999
                    ArmazenarPosicoes.idx4 = 999
                    
            End Select
        
        Case 37 'Movimento para esquerda
        
        
        Case 39 'Movimento para direita
            
        
    End Select

End Function

'FUN��ES DE USO NA DEPURA��O DO SISTEMA ********************************

Function ExibirValJogo()
'=======================================================================
'  Esta Fun��o exibe o conte�do da vari�vel "Jogo" em "txtDebug" (que
'est� em "frmJogo"
'=======================================================================

    Dim Contador1, Contador2 As Integer
    Dim Linha, Valor As String
    
    Contador1 = 1 'Indica as linhas da matriz
    Contador2 = 1 'Indica as colunas da matriz
    Linha = ""
    
    'Limpa o conte�do de "txtDebug"
    frmJogo.txtDebug.Text = ""
    
    Do While Contador1 <= 18

        Do While Contador2 <= 10
        
            Valor = CStr(Jogo(Contador1, Contador2))
            
            Select Case Len(Valor)
            
                Case 2
                    Valor = "0" & Valor
                
                Case 1
                    Valor = "00" & Valor
                
            End Select
            
            Linha = Linha & Valor & " "
            Contador2 = Contador2 + 1

        Loop

        frmJogo.txtDebug.Text = frmJogo.txtDebug.Text & Linha

        Linha = ""
        Contador2 = 1
        Contador1 = Contador1 + 1

    Loop

End Function

'***********************************************************************
