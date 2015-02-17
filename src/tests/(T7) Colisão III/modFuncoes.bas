Attribute VB_Name = "modFuncoes"
'==============================================================
'Este módulo contém as funções utilizadas no Jogo "Down Blocks"
'2005 - Desenvolvido por André Martins
'Contato: andrel_martins@yahoo.com.br
'==============================================================

Global cmdNovoJogoTEXTO(3) As String ' Em suas três posições são
                                     'armazenados os textos refe-
                                     'rentes aos rótulos do con-
                                     'trole "cmdNovoJogo", sendo:
                                     ' - 0: Iniciar
                                     ' - 1: Pausar
                                     ' - 2: Continuar

Global BlocoEmJogo(179) As Boolean ' Indica se os blocos estão
                                   'em uso em "pctJogo" ou não

Global Jogo(18, 10) As Integer ' Indica as posições disponíveis
                               'em "pctJogo", sendo a representação
                               'matemática deste container
                               ' O valor "999" indica que não há
                               'nenhum bloco nesta posição de
                               '"pctjogo". Qualquer valor diferente
                               'disso (que irá variar de "0" a "179")
                               'indica o bloco (o índice de uma das
                               'PictureBox existentes naquele container)
                               'que está nesta posição
                    
Global UmaPosicao As Integer ' Armazena o valor "375", que
                             'indica o valor mínimo para a
                             'movimentação de um bloco em
                             '"pctJogo", para qualquer direção
                                 
Global IndicesEmJogo(4) As Integer ' Armazena o índice dos blocos
                                   'que estiverem em primeiro plano
                                   'no jogo (blocos que estão aptos
                                   'a serem movidos em "pctJogo")
                                   ' O que define a quantidade de
                                   'índices armazenados é o tipo
                                   'do bloco, representado pela
                                   'variável "TipoBlocoEmJogo"
                                   
Type Posicao             ' Tipo de variável que possui duas
    PosicaoX As Integer  'possiblidadesde inserção de dados
    PosicaoY As Integer  '("PosicaoX" e "PosicaoY"), utilizado
End Type                 'para criar variáveis que armazenaram
                         'as coordenadas cartesianas X e Y do
                         'índice de um bloco na matriz "Jogo"

Type PosicoesParaIndices ' Tipo de variável que armazena índices
    idx1 As Integer      'das 4 posições dos blocos presentes
    idx2 As Integer      'Jogo. É utitilizada em diversas funções
    idx3 As Integer
    idx4 As Integer
End Type

Public Declare Function TocarSom Lib "WINMM.DLL" Alias _
 "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Function CarregarImageList(Endereco As String)
'============================================================
' Esta função carrega as imagens utilizadas nos blocos para
'um ImageList e, caso já haja dados no ImageList, deleta-os
'antes de inserir as novas imagens
'============================================================
    
    Dim Contador As Integer

    ' Se o contador do ImageList for maios que "0", indica que
    'há imagens armazenadas. Assim, exclue-as antes de iniciar
    'o carregamento das novas imagens
    If frmJogo.ImageListBlocos.ListImages.Count > 0 Then

        Contador = 7
    
        Do While Contador > 0
        
            frmJogo.ImageListBlocos.ListImages.Remove (Contador)
        
            Contador = Contador - 1
        
        Loop
    
    End If

    'Insere as novas imagens no contador, segundo a seleção do
    'usuário, indica pela variável "Endereco"
    Contador = 1

    Do While Contador <= 7
    
        frmJogo.ImageListBlocos.ListImages.Add Contador, "Bloco" & Contador, LoadPicture(Endereco & Contador & ".jpg")
    
        Contador = Contador + 1
    
    Loop

End Function

Function TrocarIdioma(Idioma As String)
'============================================================
' Esta função troca o idioma dos menus e labels dos formulá-
'rios que compõem o Jogo
'============================================================

    Dim vControle As String, vLabel As String
    Dim Menu As Object
    Dim Formulario As String

    vLinhaTXT = FreeFile
    
    ' Abre o arquivo de Idiomas (.lng), que está na pasta "\Idioma"
    'segundo a seleção do usuário (variável "Idioma"
    Open App.Path & "\Idioma\" & Idioma & ".lng" For Input As vLinhaTXT 'Abre o arquivo texto
    
    'Realiza o loop enquanto não for fim do arquivo
    Do While Not EOF(vLinhaTXT)
        
        'Lê a linha do arquivo texto onde o cursor está
        Line Input #vLinhaTXT, Linha
   
        ' Verifica se o texto da corrente linha do arquivo é
        'indicação de nome de formulário
        If Left(Linha, 1) = "*" Then
        
            Formulario = Linha
            ' Lê a próxima linha do arquivo texto (que indica
            'um controle do corrente formulário)
            Line Input #vLinhaTXT, Linha
            
        End If
   
        'Verifica de qual formulário pertence o atual componente
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
                '   - vLabel: o rótulo do controle
                '*Usa como base para separação da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))
                
                ' Verifica se o último caractere do nome do controle
                '(valor de "vControle") é um número
                If IsNumeric(Right(vControle, 1)) = True Then
        
                    ' Sendo número, modifica o rótulo do corrente objeto
                    'com o idioma selecionado mas, ao invés de usar o
                    'valor de "vControle" como base para designar o
                    'objeto, retira o número da variável e usa este
                    'como substrato para a formação do índice do corrente
                    'objeto
                    Set Menu = frmJogo(Left(vControle, Len(vControle) - 1))
                    Menu(CInt(Right(vControle, 1))).Caption = vLabel
        
                Else
                
                    'Modifica o rótulo do corrente objeto o idioma selecionado
                    Set Menu = frmJogo(vControle)
                    Menu.Caption = vLabel
        
                End If
         
            Case "*frmSobre*"
    
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o rótulo do controle
                '*Usa como base para separação da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))

                If IsNumeric(Right(vControle, 1)) = True Then
        
                    ' Sendo número, modifica o rótulo do corrente objeto
                    'com o idioma selecionado mas, ao invés de usar o
                    'valor de "vControle" como base para designar o
                    'objeto, retira o número da variável e usa este
                    'como substrato para a formação do índice do corrente
                    'objeto
                    Set Menu = frmJogo(Left(vControle, Len(vControle) - 1))
                    Menu(CInt(Right(vControle, 1))).Caption = vLabel
        
                Else
                
                    'Verifica se o controle atual é o próprio Formulário
                    If vControle = "frmSobre" Then
                
                        'Modifica o rótulo do Formulário
                        frmSobre.Caption = vLabel
                        
                    Else
                    
                        'Modifica o rótulo do corrente objeto o idioma selecionado
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
' Esta função retira os "Checks" existentes na área do menu
'que contém os Estilos dos Blocos, os Idiomas ou os Níveis do
'Jogo (segundo seleção do usuário indicada através da variável
'"Menu"
'=============================================================

    Dim Contador As Integer

    Select Case Menu
    
        Case "Estilo"
            Contador = 0
        
            'Realiza-se um laço que deverá ser executado até que se
            'atinja o número de Estilos de Blocos existentos no menu
            Do While Contador < 2
        
                'Indica "False" para o método "Checked" do controle
                frmJogo.mnuEst(Contador).Checked = False
        
                'Adiciona 1 à variável contadora
                Contador = Contador + 1
        
            Loop
            
        Case "Idioma"
            Contador = 0
        
            'Realiza-se um laço que deverá ser executado até que se
            'atinja o número de Estilos de Blocos existentos no menu
            Do While Contador < 2
        
                'Indica "False" para o método "Checked" do controle
                frmJogo.mnuIdioma(Contador).Checked = False
        
                'Adiciona 1 à variável contadora
                Contador = Contador + 1
                
            Loop
                
        Case "Nível"
            Contador = 0
        
            'Realiza-se um laço que deverá ser executado até que se
            'atinja o número de Estilos de Blocos existentos no menu
            Do While Contador < 10
        
                'Indica "False" para o método "Checked" do controle
                frmJogo.mnuNivel(Contador).Checked = False
        
                'Adiciona 1 à variável contadora
                Contador = Contador + 1
                
            Loop
        
    End Select

End Function

Function GerarBlocos(Container As String, TipoBloco As Integer, EstiloBloco As Integer)
'=============================================================
' Esta função gera os blocos aleatórios que serão utilizados
'durante o Jogo
' Há 5 tipos de Blocos que podem ser gerados, com 7 cores
'distintas para a formação destes (asteriscos representam os
'blocos):

'   - ****; representado pelo número "0"

'   - ***; representado pelo número "1"

'      *
'   - ***; representado pelo número "2"

'       *
'   - ***; representado pelo número "3"

'      **
'   - ** ; representado pelo número "4"

'     **
'   - **; representado pelo número "5"
'
' As variáveis da função indicam:
'   - Container: indica a PictureBox que conterá os Blocos
'             ("pctJogo" ou "pctProx");
'   - TipoBloco: indica o Tipo do Bloco (é um número aleatório
'             gerado através de "Random(4)"
'   - EstiloBloco: indica o Estilo do Bloco (uma imagem sele-
'             cionada aleatoriamente através de "Random(7)"
'             em ImageListBlocos
'
'=============================================================

    Dim Contador As Integer
    Dim PosX As Integer, PosY As Integer
    Dim pctObj As Object
    Dim IndexBloco As Integer 'Armazena o número do índice
    
    'Indica qual PictureBox utilizar, através do Container
    If Container = "pctJogo" Then
    
        'Armazena "999" para as posições dos índices dos blocos no Jogo
        IndicesEmJogo(1) = 999
        IndicesEmJogo(2) = 999
        IndicesEmJogo(3) = 999
        IndicesEmJogo(4) = 999
    
        'Indica a qual PictureBox se referenciará "pctObj"
        Set pctObj = frmJogo.imgBloco
    
    Else
        
        'Indica a qual PictureBox se referenciará "pctObj"
        Set pctObj = frmJogo.imgBlocoProx
    
        ' Esconde todos os Blocos utilizados em "pctProx" e
        'retorna-os a suas posições iniciais
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
                
                'Define as posições iniciais do primeiro bloco
                PosX = 3 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
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
                    
                    ' Armazena o índice deste bloco para que este
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
                
                'Define as posições iniciais do primeiro bloco
                PosX = 3 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
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

                    ' Armazena o índice deste bloco para que este
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
                
                'Define as posições iniciais do primeiro bloco
                PosX = 4 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
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
                
                ' Armazena o índice deste bloco para que este
                'seja movimentado em "pctJogo"
                IndicesEmJogo(1) = IndexBloco
                    
            End If
            
            'Prepara-se para inserir os próximos três blocos
            'Definem-se as posições iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições
                PosX = 3 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posições
                PosX = 337
                PosY = 425
                
            End If
            
            ' Contador inicia-se em 2 por já se haver inserido
            'o primeiro bloco
            Contador = 2

            'Inserem-se os próximos 3 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 2
                    
                    ' Armazena o índice deste bloco para que este
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
                
                'Define as posições iniciais do primeiro bloco
                PosX = 5 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
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
                
                ' Armazena o índice deste bloco para que este
                'seja movimentado em "pctJogo"
                IndicesEmJogo(1) = IndexBloco
                    
            End If
            
            'Prepara-se para inserir os próximos três blocos
            'Definem-se as posições iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições
                PosX = 3 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posições
                PosX = 337
                PosY = 425
                
            End If
            
            ' Contador inicia-se em 2 por já se haver inserido
            'o primeiro bloco
            Contador = 2

            'Inserem-se os próximos 3 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 3
                    
                    ' Armazena o índice deste bloco para que este
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
                
                'Define as posições iniciais do primeiro bloco
                PosX = 4 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
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
                    
                    ' Armazena o índice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If

                Contador = Contador + 1
                
            Loop
            
            'Prepara-se para inserir os próximos 2 blocos
            'Definem-se as posições iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições
                PosX = 3 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posições
                PosX = 337
                PosY = 425
                
            End If

            'Inserem-se os próximos 2 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 4
                    
                    ' Armazena o índice deste bloco para que este
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
                
                'Define as posições iniciais do primeiro bloco
                PosX = 4 * UmaPosicao
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
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
                    
                    ' Armazena o índice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If
                
                Contador = Contador + 1
                
            Loop
            
            'Prepara-se para inserir os próximos 2 blocos
            'Definem-se as posições iniciais do primeiro bloco
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições
                PosX = 4 * UmaPosicao
                PosY = UmaPosicao
            
            Else
            
                'Define as posições
                PosX = 525
                PosY = 425
                
            End If

            'Inserem-se os próximos 2 blocos
            Do While Contador <= 4
           
                IndexBloco = ProxIndexBloco(Container)
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + UmaPosicao
                
                If Container = "pctJogo" Then

                    InserirNaMatriz IndexBloco, Contador, 5
                    
                    ' Armazena o índice deste bloco para que este
                    'seja movimentado em "pctJogo"
                    IndicesEmJogo(Contador) = IndexBloco
                    
                End If
                
                Contador = Contador + 1
                
            Loop

    End Select

'USO NA DEPURAÇÃO DO SISTEMA ********************************
    
    'Atualiza a exibição da matriz "Jogo" em "txtDebug"
    ExibirValJogo
    
'************************************************************
    
End Function

Function Random(Base_de_Calculo As Integer, TipoDeAleatoriedade As Integer)
'=======================================================================
'  Esta Função, com base nos segundos do relógio interno do computador,
' seleciona uma das três possibilidades de cálculo de aleatoriedade e
' realiza a seleção de um número aleatório
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
'  Esta Função indica qual é o próximo índice (que se referenciará a um
'bloco durante o jogo) a ser utilizado
'  A variável "ContainerDoBloco" indica qual é o container de onde os blocos
'deverão ser selecionados ("pctJogo" ou "pctProx"), uma vez que os
'blocos destes seguem numerações diferentes.

'  Quando o container for "pctJogo":
'  A função utiliza os dados da variável "BlocoEmJogo", que armazena
'"True" para blocos que estão em uso no Jogo (inseridos no container
'"pctJogo") e "False" para blocos que não estão em uso. Através da
'indicação da variável, a função busca o próximo bloco disponível
'para utilização a partir do índice "0", selecionando como bloco
'disponível aquele que tem como dado armazenado "False"

'  Quando o container for "pctProx":
'  Verificar-se-á a posição dos blocos de "pctProx" e aquele que ainda
'estiver em sua posição incial "Top = 0" e "Left = 0" será utilizado
'=======================================================================
    
    Dim Contador As Integer
    
    Select Case ContainerDoBloco
    
        Case "pctJogo"
    
            Contador = 0
            
            Do While Contador < 180
            
                If BlocoEmJogo(Contador) = False Then
                
                    ProxIndexBloco = Contador
                    
                    'Indica que este bloco está em uso no Jogo
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
'  Esta Função insere as posições do bloco com "Indice" e "TipoBloco"
'selecionados na matriz "Jogo" (válido quando da inserção dos blocos
'em "pctJogo"), levando-se em conta a "Posicao" de cada bloco no conjunto.
'=========================================================================

'  Obs.: Os números após as representações dos blocos nas descrições do CASE
'abaixo indicam a "Posicao" de cada bloco no conjunto

    Select Case Tipo_do_Bloco
    
        Case 0
            'Insere um conjunto de Blocos do tipo **** ( 1234 ) na matriz "Jogo"
            'EIXO DO BLOCO: posição "2"
            
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
            'EIXO DO BLOCO: posição "2"
            
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
            'EIXO DO BLOCO: posição "3"
            
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
            'EIXO DO BLOCO: posição "3"
            
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
            'EIXO DO BLOCO: posição "4"
            
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
            'EIXO DO BLOCO: <não possui eixo (não há rotação)>
            
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
'  Esta Função localiza a posição (X,Y) de um bloco na matriz, através da
'indicação de seu índice ("Indice_do_Bloco")
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
'  Esta Função verifica cada linha na matriz, observando se todas as
'posições de coluna (eixo X) estão com valores diferentes de "999", o que
'indica uma linha completa, e retorna o número desta linha
' Obs.: inicia-se a vericação a partir da última linha; retorna-se "999"
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
'  Esta Função reordena os blocos e os valores a estes relacionados na
'matriz "Jogo" através de "Linha_da_Matriz", que indica a linha que deverá
'ser excluída (e os blocos acima desta movidos para esta posição)
'=========================================================================

    Dim Contador_da_Reordem As Integer
    Dim Contador_das_Colunas As Integer

    'Inicia um DO que executar-se-á durante "Linha_da_Matriz - 1" vezes
    Contador_da_Reordem = (Linha_da_Matriz - 1)
    Contador_das_Colunas = 1
    
    Do While Contador_da_Reordem >= 1
    
        Do While Contador_das_Colunas <= 10
        
            'Verifica se há bloco nesta posição
            If Jogo(Contador_da_Reordem, 1) <> 999 Then
            
                'Move o bloco uma posição abaixo
                frmJogo.imgBloco(Jogo(Contador_da_Reordem, 1)).Top = Contador_da_Reordem * UmaPosicao
                'Indica a nova posição do bloco na matrtiz
                Jogo(Linha_da_Matriz, 1) = Jogo(Contador_da_Reordem, 1)
                'Armazena "999" na antiga posição do bloco na matriz
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
'  Esta Função move os blocos de índices "I1", "I2", "I3" e "I4" para:
'       - Baixo ("Direcao = 40")
'       - Esquerda ("Direcao = 37")
'       - Direita ("Direcao = 39")
'=========================================================================

    Dim Blocos_a_mover(4) As Integer
    Dim Contador_MoverBloco As Integer
    Dim Posicao_Bloco_Matriz As Posicao
    'Armazenam as variações de posições para a manipulação na matriz "Jogo"
    Dim PosicaoMatriz_X As Integer, PosicaoMatriz_Y As Integer
    'Armazenam as variações de posições para a manipulação dos blocos
    Dim PosicaoBloco_X As Integer, PosicaoBloco_Y As Integer

    Select Case Direcao
    
        Case 40
            'Move os blocos uma posição para baixo
     
            PosicaoMatriz_X = 0
            PosicaoMatriz_Y = 1
            PosicaoBloco_X = 0
            PosicaoBloco_Y = UmaPosicao
            
            ' Ao mover-se para a direita, como há acréscimo
            'de UMA posição, deve-se mover primeiramente o
            'ÚLTIMO bloco do conjunto, e depois os demais,
            'em ordem decrescente
            Blocos_a_mover(1) = I1
            Blocos_a_mover(2) = I2
            Blocos_a_mover(3) = I3
            Blocos_a_mover(4) = I4
        
        Case 37
            'Move os blocos uma posição para a esquerda
        
            PosicaoMatriz_X = -1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = -UmaPosicao
            PosicaoBloco_Y = 0
            
            ' Ao mover-se para a esquerda, como há diminuição
            'de UMA posição, deve-se mover primeiramente o
            'PRIMEIRO bloco do conjunto, e depois os demais,
            'em ordem crescente
            Blocos_a_mover(1) = I4
            Blocos_a_mover(2) = I3
            Blocos_a_mover(3) = I2
            Blocos_a_mover(4) = I1
        
        Case 39
            'Move os blocos uma posição para a direita
            
            PosicaoMatriz_X = 1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = UmaPosicao
            PosicaoBloco_Y = 0
            
            ' Ao mover-se para a baixo, como há acréscimo
            'de UMA posição, deve-se mover primeiramente o
            'ÚLTIMO bloco do conjunto, e depois os demais,
            'em ordem decrescente
            Blocos_a_mover(1) = I1
            Blocos_a_mover(2) = I2
            Blocos_a_mover(3) = I3
            Blocos_a_mover(4) = I4
            
    End Select
    
    Contador_MoverBloco = 4
        
    Do While Contador_MoverBloco >= 1
            
        If Blocos_a_mover(Contador_MoverBloco) <> 999 Then
    
            'Move o bloco uma posição abaixo na matriz
            Posicao_Bloco_Matriz = LocalizarBlocoNaMatriz(Blocos_a_mover(Contador_MoverBloco))
            Jogo(Posicao_Bloco_Matriz.PosicaoY + PosicaoMatriz_Y, Posicao_Bloco_Matriz.PosicaoX + PosicaoMatriz_X) = Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX)
            Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX) = 999
                    
            'Move o bloco 1 posição abaixo em "pctJogo"
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top + PosicaoBloco_Y)
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left + PosicaoBloco_X)
                
        End If
            
        Contador_MoverBloco = Contador_MoverBloco - 1
            
    Loop

End Function

Function ArmazenarPosicoes(Direcao_Movimento As Integer, Tipo_BlocoEmJogo As Integer, _
PosicaoAtual_BlocoEmJogo As Integer) As PosicoesParaIndices
'=========================================================================
'  Esta Função indica quais posições de índices de um determinado
'"Tipo_BlocoEmJogo" devem ser verificadas, usando, para tal, os dados da
'"Direcao_Movimento" que deve ser analisada e da "PosicaoAtual_BlocoEmJogo"
'=========================================================================

    'Primeiramente, verifica qual a direção que deve ser analisada
    Select Case Direcao_Movimento
    
        Case 40 'Descida
        
            ' Verifica qual o tipo de Bloco
            Select Case Tipo_BlocoEmJogo
            
                Case 0
                    'Blocos do tipo ****
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica todas as posições abaixo do
                            'conjunto de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 4
                        
                        Case 1
                            ' Verifica apenas a posição diretamente
                            'abaixo do bloco "4" do conjunto
                            ArmazenarPosicoes.idx1 = 4
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 1
                    'Blocos do tipo ***
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica todas as posições abaixo do
                            'conjunto de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas a posição diretamente
                            'abaixo do bloco "3" do conjunto
                            ArmazenarPosicoes.idx1 = 3
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 2
                    '                *
                    'Blocos do tipo ***
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "1", "2" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 3
                    '                 *
                    'Blocos do tipo ***
        
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "1", "2" e "3"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posições diretamente
                            'abaixo dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 4
                    '                **
                    'Blocos do tipo **
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica apenas as posições diretamente
                            'abaixo dos blocos "3" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 3
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posições diretamente
                            'abaixo dos blocos "2" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 5
                    '               **
                    'Blocos do tipo **
                    
                    ' Neste tipo de bloco, apenas as posições "3" e "4"
                    'são sempre verificadas
                    ArmazenarPosicoes.idx1 = 3
                    ArmazenarPosicoes.idx2 = 4
                    ArmazenarPosicoes.idx3 = 999
                    ArmazenarPosicoes.idx4 = 999
                    
            End Select
        
        Case 37 'Movimento para esquerda
        
        
        Case 39 'Movimento para direita
            
        
    End Select

End Function

'FUNÇÕES DE USO NA DEPURAÇÃO DO SISTEMA ********************************

Function ExibirValJogo()
'=======================================================================
'  Esta Função exibe o conteúdo da variável "Jogo" em "txtDebug" (que
'está em "frmJogo"
'=======================================================================

    Dim Contador1, Contador2 As Integer
    Dim Linha, Valor As String
    
    Contador1 = 1 'Indica as linhas da matriz
    Contador2 = 1 'Indica as colunas da matriz
    Linha = ""
    
    'Limpa o conteúdo de "txtDebug"
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
