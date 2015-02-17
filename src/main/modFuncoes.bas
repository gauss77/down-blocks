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
                                     
Global cmdNovoJogoSTATUS As Integer ' Indica qual a posição da va-
                                    'riável "cmdNovoJogoTEXTO"
                                    'está sendo usada no momento
                                    'como rótulo em "cmdNovoJogo"
                                     
Global MsgBoxGravarRecorde As String ' Armazena o texto a ser
                                     'utilizado na MessageBox
                                     'ao indicar falha na
                                     'confecção do nome do jogador
                                     'durante a inserção de um
                                     'novo Recorde

Global BlocoEmJogo(179) As Boolean ' Indica se os blocos estão
                                   'em uso em "pctJogo" ou não

Global EstiloBlocoEmJogo(179) As Integer ' Indica os estilos (número de
                                   '1 a 7) de cada bloco existente
                                   'no jogo

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
                                   
Global RecordesJogo(10, 2) As String ' Armazena os 10 primeiros recordes
                                     'do Jogo, juntamente com o nome do
                                     'jogador
                                 
Global ConfigJogo(7, 2) As String ' Armazena as configurações presentes
                                  '"config.ini" para posterior manipulação

Global EstiloBlocos As Integer ' Armazena o Estilo dos blocos que estão em
                               'Jogo
                               
Global IdiomaJogo As String 'Armazena o Idioma do Jogo

Global NivelJogo As Integer 'Armazena o Nível do Jogo

Global MusicaExecutando As Boolean ' Indica se Músicas MIDI estão
                                   'sendo executadas ou não
                        
Global TempoMusica As Integer 'Armazena o Tempo Inicial da Música em execução
                        
                        
                        
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
                 
Public Enum SndPlayFlags ' Enumeração utilizada na função
  SND_SYNC = &H0         'para tocar arquivos WAVE
  SND_ASYNC = &H1
  SND_NODEFAULT = &H2
  SND_MEMORY = &H4
  SND_LOOP = &H8
  SND_NOSTOP = &H10
End Enum

Public Enum AcaoTocador ' Enumeração utilizada na função
  Tocar                 'para tocar arquivos MIDI
  Pausar
  Resumir
  Parar
  Tempo
End Enum



'Função para tocar arquivos WAVE
Public Declare Function sndPlaySound Lib "winmm.dll" _
       Alias "sndPlaySoundA" (ByVal lpszSoundName As _
       String, ByVal uFlags As Long) As Long
       
'Função para tocar arquivos MIDI
Public Declare Function mciExecute Lib "winmm.dll" _
       (ByVal lpstrCommand As String) As Long
       
'Função para uso da conversão de caminhos longos de arquivos
'Extraída de FreeVBCode.com (http://www.freevbcode.com)
Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
    As Long
   
   
   
Function RetornarCaminhoDOS(ByVal Caminho As String) As String
'==============================================================
' Esta função retorna um caminho de arquivo no formato 8.3
'(DOS) dado este no formato atual de 255 caracteres, retornando
'"" caso o caminho não exista.
'
' Extraída de FreeVBCode.com (http://www.freevbcode.com)
'==============================================================
    
    Dim lAns As Long
    Dim sAns As String
    Dim iLen As Integer
       
    On Error Resume Next
    
    If Dir(Caminho) = "" Then Exit Function
    
    sAns = Space(255)
    lAns = GetShortPathName(Caminho, sAns, 255)
    RetornarCaminhoDOS = Left(sAns, lAns)
        
End Function
       
Function TocarSom(ArquivoWAVE As String, Optional Flags As SndPlayFlags)
'============================================================
' Esta função toca um arquivo WAVE
'
' Opções do parâmetro "Flags":
' - "SND_SYNC" - Toca o WAVE sincronizado (default).
' - "SND_ASYNC" - Toca o WAVE sem sincronismo.
' - "SND_NODEFAULT" - Não usa o som padrão.
' - "SND_MEMORY" - Direciona o "IpszSoundName" para o lugar
'    de um arquivo na memória.
' - "SND_LOOP" - Toca o WAVE em looping, até que se o mande
'    parar
' - "SND_NOSTOP" - Não pára o som que estiver tocando.
'
' Obs.:
'   - Para tocar o som WAVE, utilize:
'       TocarSom "C:\<Nome Arquivo>.WAV"
'   - Para parar o som WAVE, utilize:
'       TocarSom ""
'============================================================

    If ArquivoWAVE = "" Then

        sndPlaySound 0&, 0

    Else
    
        sndPlaySound ArquivoWAVE, Flags

  End If

End Function

Function TocarMusica(Acao As AcaoTocador, Optional ArquivoMIDI As String, _
    Optional Tempo As Integer)
'===============================================================
' Esta função coordena o tocador de MIDI, utilizando-se para
'tal os parâmetros presentes na função.
' Parâmetros:
'   - "Acao": Baseado na enumeração "AcaoTocador", possui as
'     seguintes opções:
'          * "Tocar" - Abre uma música MIDI, usando como base
'            o parâmetro "Arquivo
'          * "Pausar" - Pausa uma música MIDI em execução
'          * "Resumir" - Reinicia uma música MIDI anteriormente
'            pausada
'          * "Parar" - Fecha o arquivo MIDI
'          * "Tempo" - aumenta ou diminui o Tempo da
'            música MIDIcom base no parâmetro "Tempo"
'
' Obs.: "MusicaMIDI" é o nome dado ao processo que aciona o
'arquivo MIDI
'===============================================================
     
    Select Case Acao
    
        Case 0
            mciExecute "open " & RetornarCaminhoDOS(ArquivoMIDI) & " type sequencer alias Musica"
            
            mciExecute "play Musica"
            
            'Indica que uma música está sendo tocada
            MusicaExecutando = True
        
        Case 1
            mciExecute "pause Musica"
        
        Case 2
            mciExecute "play Musica"
        
        Case 3
            'pergunta se há uma música em execução para fechá-la
            If MusicaExecutando = True Then
            
                mciExecute "close Musica"
                
                MusicaExecutando = False
            
            End If
        
        Case 4
            mciExecute "set Musica tempo " & CStr(Tempo)
    
    End Select
   
End Function

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
    Dim Objeto_a_Manipular As Object
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
                Select Case vControle
                
                    Case "cmdNovoJogoTEXTO0"
                    
                        cmdNovoJogoTEXTO(0) = vLabel
        
                    Case "cmdNovoJogoTEXTO1"
                    
                        cmdNovoJogoTEXTO(1) = vLabel
                    
                    Case "cmdNovoJogoTEXTO2"
                    
                        cmdNovoJogoTEXTO(2) = vLabel
        
                    Case "MsgBoxGravarRecorde"
                    
                        MsgBoxGravarRecorde = vLabel
                
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
                    Set Objeto_a_Manipular = frmJogo(Left(vControle, Len(vControle) - 1))
                    Objeto_a_Manipular(CInt(Right(vControle, 1))).Caption = vLabel
        
                Else
                
                    'Modifica o rótulo do corrente objeto o idioma selecionado
                    Set Objeto_a_Manipular = frmJogo(vControle)
                    Objeto_a_Manipular.Caption = vLabel
        
                End If
         
            Case "*frmSobre*"
    
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o rótulo do controle
                '*Usa como base para separação da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))

                'Verifica se o controle atual é o próprio Formulário
                If vControle = "frmSobre" Then
                
                    'Modifica o rótulo do Formulário
                    frmSobre.Caption = vLabel
                        
                Else
                    
                    'Modifica o rótulo do corrente objeto o idioma selecionado
                    Set Objeto_a_Manipular = frmSobre(vControle)
                    Objeto_a_Manipular.Caption = vLabel
        
                End If
                
            Case "*frmSalvarConfig*"
    
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o rótulo do controle
                '*Usa como base para separação da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))

                'Modifica o rótulo do corrente objeto o idioma selecionado
                Set Objeto_a_Manipular = frmSalvarConfig(vControle)
                Objeto_a_Manipular.Caption = vLabel
                       
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
           
                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco
                    
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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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
    
    Dim ContadorProxIndex As Integer
    
    Select Case ContainerDoBloco
    
        Case "pctJogo"
    
            ContadorProxIndex = 0
            
            Do While ContadorProxIndex < 180
            
                If BlocoEmJogo(ContadorProxIndex) = False Then
                
                    ProxIndexBloco = ContadorProxIndex
                    
                    'Indica que este bloco está em uso no Jogo
                    BlocoEmJogo(ContadorProxIndex) = True
                    Exit Do
                
                End If
                
                ContadorProxIndex = ContadorProxIndex + 1
            
            Loop
            
        Case "pctProx"
        
            ContadorProxIndex = 0
            
            Do While ContadorProxIndex < 4

                If frmJogo.imgBlocoProx(ContadorProxIndex).Top = 0 Then
                
                    ProxIndexBloco = ContadorProxIndex
                    Exit Do
                
                End If
                
                ContadorProxIndex = ContadorProxIndex + 1
            
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
    
    Do While Contador_da_Reordem >= 1
    
        Contador_das_Colunas = 1
    
        Do While Contador_das_Colunas <= 10
        
            'Verifica se há bloco nesta posição
            If Jogo(Contador_da_Reordem, Contador_das_Colunas) <> 999 Then
            
                'Move o bloco uma posição abaixo
                frmJogo.imgBloco(Jogo(Contador_da_Reordem, Contador_das_Colunas)).Top = (Contador_da_Reordem) * UmaPosicao
                'Indica a nova posição do bloco na matrtiz
                Jogo(Contador_da_Reordem + 1, Contador_das_Colunas) = Jogo(Contador_da_Reordem, Contador_das_Colunas)
                'Armazena "999" na antiga posição do bloco na matriz
                Jogo(Contador_da_Reordem, Contador_das_Colunas) = 999
            
            End If
        
            Contador_das_Colunas = Contador_das_Colunas + 1
        
        Loop
    
        Contador_das_Colunas = 1
        Contador_da_Reordem = Contador_da_Reordem - 1
    
    Loop

End Function

Function MoverBloco(Direcao As Integer, I1 As Integer, I2 As Integer, _
 I3 As Integer, I4 As Integer, TipoBloco_a_Mover As Integer, _
 PosicaoDoBloco_a_Mover As Integer)
'============================================================================
'  Esta Função move os blocos de índices "I1", "I2", "I3" e "I4" para:
'       - Cima ("Direcao = 38")
'       - Baixo ("Direcao = 40")
'       - Esquerda ("Direcao = 37")
'       - Direita ("Direcao = 39")
'  Como a Função começa sempre o movimento do Bloco a partir da posição 1
'de  "Blocos_a_mover()", há a verificação de quais posições devem ser
'movidas primeiro (através de "TipoBloco_a_Mover" e "PosicaoDoBloco_a_Mover")
'para evitar sobreposição de  blocos.
'============================================================================

    Dim Blocos_a_mover(4) As Integer
    Dim Contador_MoverBloco As Integer
    Dim Posicao_Bloco_Matriz As Posicao
    'Armazenam as variações de posições para a manipulação na matriz "Jogo"
    Dim PosicaoMatriz_X As Integer, PosicaoMatriz_Y As Integer
    'Armazenam as variações de posições para a manipulação dos blocos
    Dim PosicaoBloco_X As Integer, PosicaoBloco_Y As Integer

    Select Case Direcao
    
        Case 38
        'Move os blocos uma posição para cima
     
            PosicaoMatriz_X = 0
            PosicaoMatriz_Y = -1
            PosicaoBloco_X = 0
            PosicaoBloco_Y = -UmaPosicao
                        
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                    End Select
                            
                Case 1
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
          
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                            
                    End Select
                        
                Case 2
                '                *
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I4
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I2
                                    
                        Case 2
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 3
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                    End Select
                        
                Case 3
                '                 *
                'Blocos do tipo ***
                
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I4
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I2
                                    
                        Case 2
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 3
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                    End Select
                        
                Case 4
                '                **
                'Blocos do tipo **
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                    End Select
                            
                Case 5
                '               **
                'Blocos do tipo **
                    
                    Blocos_a_mover(1) = I1
                    Blocos_a_mover(2) = I2
                    Blocos_a_mover(3) = I3
                    Blocos_a_mover(4) = I4
                    
            End Select
    
        Case 40
        'Move os blocos uma posição para baixo
     
            PosicaoMatriz_X = 0
            PosicaoMatriz_Y = 1
            PosicaoBloco_X = 0
            PosicaoBloco_Y = UmaPosicao
            
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                    End Select
                            
                Case 1
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
          
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                            
                    End Select
                        
                Case 2
                '                *
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                    
                        Case 2
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 3
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I4
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I2
                                
                    End Select
                        
                Case 3
                '                 *
                'Blocos do tipo ***
                
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                    
                        Case 2
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 3
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I4
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I2
                                
                    End Select
                        
                Case 4
                '                **
                'Blocos do tipo **
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                    End Select
                            
                Case 5
                '               **
                'Blocos do tipo **
                    
                    Blocos_a_mover(1) = I4
                    Blocos_a_mover(2) = I3
                    Blocos_a_mover(3) = I2
                    Blocos_a_mover(4) = I1
                    
            End Select
        
        Case 37
        'Move os blocos uma posição para a esquerda
        
            PosicaoMatriz_X = -1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = -UmaPosicao
            PosicaoBloco_Y = 0
            
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    Blocos_a_mover(1) = I1
                    Blocos_a_mover(2) = I2
                    Blocos_a_mover(3) = I3
                    Blocos_a_mover(4) = I4
                            
                Case 1
                'Blocos do tipo ***
                            
                    Blocos_a_mover(1) = I1
                    Blocos_a_mover(2) = I2
                    Blocos_a_mover(3) = I3
                    Blocos_a_mover(4) = I4
                        
                Case 2
                '                *
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                    
                        Case 2
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 3
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                    End Select
                        
                Case 3
                '                 *
                'Blocos do tipo ***
                
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 1
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                    
                        Case 2
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 3
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                    End Select
                        
                Case 4
                '                **
                'Blocos do tipo **

                    Blocos_a_mover(1) = I1
                    Blocos_a_mover(2) = I2
                    Blocos_a_mover(3) = I3
                    Blocos_a_mover(4) = I4
                            
                Case 5
                '               **
                'Blocos do tipo **
                    
                    Blocos_a_mover(1) = I1
                    Blocos_a_mover(2) = I2
                    Blocos_a_mover(3) = I3
                    Blocos_a_mover(4) = I4
        
            End Select
        
        Case 39
        'Move os blocos uma posição para a direita
            
            PosicaoMatriz_X = 1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = UmaPosicao
            PosicaoBloco_Y = 0
            
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                    End Select
                    
                Case 1
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                    End Select
                        
                Case 2
                '                *
                'Blocos do tipo ***
                            
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                    
                        Case 2
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 3
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                    End Select
                        
                Case 3
                '                 *
                'Blocos do tipo ***
                
                    'Verifica a posição atual do bloco
                    Select Case PosicaoDoBloco_a_Mover
                            
                        Case 0
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                
                        Case 1
                            Blocos_a_mover(1) = I4
                            Blocos_a_mover(2) = I3
                            Blocos_a_mover(3) = I2
                            Blocos_a_mover(4) = I1
                                    
                        Case 2
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                        Case 3
                            Blocos_a_mover(1) = I1
                            Blocos_a_mover(2) = I2
                            Blocos_a_mover(3) = I3
                            Blocos_a_mover(4) = I4
                                
                    End Select
                        
                Case 4
                '                **
                'Blocos do tipo **
                            
                    Blocos_a_mover(1) = I4
                    Blocos_a_mover(2) = I3
                    Blocos_a_mover(3) = I2
                    Blocos_a_mover(4) = I1
                            
                Case 5
                '               **
                'Blocos do tipo **
                    
                    Blocos_a_mover(1) = I4
                    Blocos_a_mover(2) = I3
                    Blocos_a_mover(3) = I2
                    Blocos_a_mover(4) = I1
             
            End Select
            
    End Select
    
    Contador_MoverBloco = 1
        
    Do While Contador_MoverBloco <= 4
            
        If Blocos_a_mover(Contador_MoverBloco) <> 999 Then
    
            'Move o bloco uma posição acima/abaixo/esquerda/direita na matriz
            Posicao_Bloco_Matriz = LocalizarBlocoNaMatriz(Blocos_a_mover(Contador_MoverBloco))
            Jogo(Posicao_Bloco_Matriz.PosicaoY + PosicaoMatriz_Y, Posicao_Bloco_Matriz.PosicaoX + PosicaoMatriz_X) = Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX)
            Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX) = 999
                    
            'Move o bloco 1 posição acima/abaixo/esquerda/direita em "pctJogo"
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top + PosicaoBloco_Y)
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left + PosicaoBloco_X)
                
        End If
            
        Contador_MoverBloco = Contador_MoverBloco + 1
            
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
        
            'Verifica qual o tipo de Bloco
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
                            ArmazenarPosicoes.idx3 = 4
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
                            'abaixo dos blocos "2", "3" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posições diretamente
                            'abaixo dos blocos "1" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 1
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
        
            'Verifica qual o tipo de Bloco
            Select Case Tipo_BlocoEmJogo
            
                Case 0
                    'Blocos do tipo ****
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posição "1"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posições do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 4
                        
                    End Select
                    
                Case 1
                    'Blocos do tipo ***
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posição "1"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posições do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 2
                    '                *
                    'Blocos do tipo ***
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "1", "2" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 3
                    '                 *
                    'Blocos do tipo ***
        
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "1", "2" e "3"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posições diretamente
                            'à esquerda dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 4
                    '                **
                    'Blocos do tipo **
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica apenas as posições diretamente
                            'à esquerda dos blocos "1" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posições diretamente
                            'à esquerda dos blocos "1", "2" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 5
                    '               **
                    'Blocos do tipo **
                    
                    ' Neste tipo de bloco, apenas as posições "1" e "3"
                    'são sempre verificadas
                    ArmazenarPosicoes.idx1 = 1
                    ArmazenarPosicoes.idx2 = 3
                    ArmazenarPosicoes.idx3 = 999
                    ArmazenarPosicoes.idx4 = 999
                    
            End Select
        
        Case 39 'Movimento para direita
            
            'Verifica qual o tipo de Bloco
            Select Case Tipo_BlocoEmJogo
            
                Case 0
                    'Blocos do tipo ****
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posição "4"
                            ArmazenarPosicoes.idx1 = 4
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posições do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 4
                        
                    End Select
                    
                Case 1
                    'Blocos do tipo ***
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posição "3"
                            ArmazenarPosicoes.idx1 = 3
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posições do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 2
                    '                *
                    'Blocos do tipo ***
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posições diretamente
                            'à direita dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posições diretamente
                            'à direita dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posições diretamente
                            'à direita dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posições diretamente
                            'à direita dos blocos "1", "2" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 3
                    '                 *
                    'Blocos do tipo ***
        
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posições diretamente
                            'à direita dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posições diretamente
                            'à direita dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posições diretamente
                            'à direita dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posições diretamente
                            'à direita dos blocos "1", "2" e "3"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 4
                    '                **
                    'Blocos do tipo **
                    
                    'Verifica a posição atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica apenas as posições diretamente
                            'à direita dos blocos "2" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posições diretamente
                            'à direita dos blocos "1", "3" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 5
                    '               **
                    'Blocos do tipo **
                    
                    ' Neste tipo de bloco, apenas as posições "2" e "4"
                    'são sempre verificadas
                    ArmazenarPosicoes.idx1 = 2
                    ArmazenarPosicoes.idx2 = 4
                    ArmazenarPosicoes.idx3 = 999
                    ArmazenarPosicoes.idx4 = 999
                    
            End Select
        
    End Select

End Function

Function MenuNivel(Acao As String)
'=======================================================================
'  Esta Função habilita (Acao = "Habilitar") os menus referentes à
'seleção de Nível de Jogo ou desabilita-os (Acao = "Desabilitar")
'=======================================================================

    Dim ContadorNivel As Integer
    
    ContadorNivel = 0

    Select Case Acao
    
        Case "Habilitar"

            'Habilita os menus selecionadores de Nível de Jogo
            Do While ContadorNivel <= 9
                
                frmJogo.mnuNivel(ContadorNivel).Enabled = True
                
                ContadorNivel = ContadorNivel + 1
                
            Loop
    
        Case "Desabilitar"
        
            'Desabilita os menus selecionadores de Nível de Jogo
            Do While ContadorNivel <= 9
                
                frmJogo.mnuNivel(ContadorNivel).Enabled = False
                
                ContadorNivel = ContadorNivel + 1
                
            Loop
        
        
    End Select

End Function

Function SelecionarNivel(Nivel_do_Jogo As Integer, EmJogo As Boolean)
'=======================================================================
'  Esta Função seleciona o nível do jogo com base no número informado
'através de "Nivel_do_Jogo" alterando, inclusive, as indicações de
'nível selecionado no menu e em "lblPontos".
'Obs.: a informação de nível nos menus só será alterada se
'"EmJogo" = "False" (indicando que a  mudança de nível ocorre por
'intervenção do usuário e não da Engine do Jogo
'=======================================================================

    Select Case Nivel_do_Jogo
    
        Case 0
            frmJogo.lblNivel.Caption = 0
            frmJogo.TimerJogo.Interval = 1000
            
            If EmJogo = False Then frmJogo.mnuNivel(0).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica
            
        Case 1
            frmJogo.lblNivel.Caption = 1
            frmJogo.TimerJogo.Interval = 880
            
            If EmJogo = False Then frmJogo.mnuNivel(1).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 10
            
        Case 2
            frmJogo.lblNivel.Caption = 2
            frmJogo.TimerJogo.Interval = 760
            
            If EmJogo = False Then frmJogo.mnuNivel(2).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 15
            
        Case 3
            frmJogo.lblNivel.Caption = 3
            frmJogo.TimerJogo.Interval = 640
            
            If EmJogo = False Then frmJogo.mnuNivel(3).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 30
            
        Case 4
            frmJogo.lblNivel.Caption = 4
            frmJogo.TimerJogo.Interval = 520
            
            If EmJogo = False Then frmJogo.mnuNivel(4).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 50
            
        Case 5
            frmJogo.lblNivel.Caption = 5
            frmJogo.TimerJogo.Interval = 400
            
            If EmJogo = False Then frmJogo.mnuNivel(5).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 70
            
        Case 6
            frmJogo.lblNivel.Caption = 6
            frmJogo.TimerJogo.Interval = 280
            
            If EmJogo = False Then frmJogo.mnuNivel(6).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 90
            
        Case 7
            frmJogo.lblNivel.Caption = 7
            frmJogo.TimerJogo.Interval = 160
            
            If EmJogo = False Then frmJogo.mnuNivel(7).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 110
            
        Case 8
            frmJogo.lblNivel.Caption = 8
            frmJogo.TimerJogo.Interval = 40
            
            If EmJogo = False Then frmJogo.mnuNivel(8).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 130
            
        Case 9
            frmJogo.lblNivel.Caption = 9
            frmJogo.TimerJogo.Interval = 1
            
            If EmJogo = False Then frmJogo.mnuNivel(9).Checked = True
            
            If MusicaExecutando = True Then TocarMusica Tempo, , TempoMusica + 150
            
    End Select

End Function

Function DetectarColisao(Indice_Bloco As Integer, _
  Direcao_a_Verificar As Integer, Optional Posicao_a_Verificar _
  As Integer) As Boolean
'=========================================================================
'  Esta Função detecta a colisão, em "pctJogo", do bloco de "Indice_Bloco"
'selecionado analisando a "Direcao_a_Verificar" escolhida
'  O parâmetro "Posicao_a_Verificar" indica quantas posições abaixo do
'bloco de "Indice_Bloco" selecionado devem ser puladas para
'a verificação. O padrão é "1" (indica que a posição imediatamente
'acima/abaixo/à esquerda/à direita do bloco deve ser verificada;
'caso fosse "2", por exemplo, a função desconsideraria a posição
'imediatamente acima/abaixo/à esquerda/à direita e analisaria a
'próxima posição após esta)
'=========================================================================

    Dim Posicao_na_Matriz As Posicao

    'Caso "Posicao_a_Verificar" seja igual a "0" (parâmetro
    'não informado), esta passa a valer "1" (valor padrão)
    If Posicao_a_Verificar = 0 Then
    
        Posicao_a_Verificar = 1
        
    End If

    'A princípio, não há colisão do bloco
    DetectarColisao = False

    'Detecta a posição do bloco de "Indice_Bloco" selecionado
    Posicao_na_Matriz = LocalizarBlocoNaMatriz(Indice_Bloco)

    'Detecta a colisão na direção selecionada
    Select Case Direcao_a_Verificar
    
        Case 38 'Direção para cima
            
            ' Verifica se é a primeira linha ("PosicaoY = 1")
            'em que o bloco está
            If Posicao_na_Matriz.PosicaoY = 1 Then
        
                ' Sendo, indica-se colisão (uma vez que não há como
                'o Bloco subir
                DetectarColisao = True
            
            Else
            ' Não havendo colisão com "pctJogo", verifica-se
            'se hã colisão com a posição imediatamente acima
            'do bloco selecionado
            
                If Posicao_na_Matriz.PosicaoY - Posicao_a_Verificar >= 1 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY - Posicao_a_Verificar, Posicao_na_Matriz.PosicaoX) <> 999 Then
                        
                        'Indica colisão
                        DetectarColisao = True
                    
                    End If
                
                Else
                
                    'Indica colisão
                    DetectarColisao = True
                
                End If
            
            End If
    
        Case 40 'Direção para baixo
            
            ' Verifica se é a última linha ("PosicaoY = 18")
            'em que o bloco está
            If Posicao_na_Matriz.PosicaoY = 18 Then
        
                ' Sendo, indica-se colisão (para que o movimento do
                'bloco cesse
                DetectarColisao = True
            
            Else
            ' Não havendo colisão com "pctJogo", verifica-se
            'se hã colisão com a posição imediatamente abaixo
            'do bloco selecionado
            
                If Posicao_na_Matriz.PosicaoY + Posicao_a_Verificar <= 18 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY + Posicao_a_Verificar, Posicao_na_Matriz.PosicaoX) <> 999 Then
                        
                        'Indica colisão
                        DetectarColisao = True
                    
                    End If
                
                Else
                
                    'Indica colisão
                    DetectarColisao = True
                
                End If
            
            End If
    
        Case 39 'Direção para o lado direito
            
            ' Verifica se é a última posição possível para a
            'direita ("PosicaoX = 10") em que o bloco está
            If Posicao_na_Matriz.PosicaoX = 10 Then
            
                ' Sendo, indica-se colisão (para que o movimento do
                'bloco cesse
                DetectarColisao = True
            
            Else
            ' Não havendo colisão com "pctJogo", verifica-se
            'se hã colisão com a posição imediatamente à direita
            'do bloco selecionado
                
                If Posicao_na_Matriz.PosicaoX + Posicao_a_Verificar <= 10 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY, Posicao_na_Matriz.PosicaoX + Posicao_a_Verificar) <> 999 Then
                        
                        'Indica colisão
                        DetectarColisao = True
                    
                    End If
                
                Else
                
                    'Indica colisão
                    DetectarColisao = True
                
                End If
            
            End If

        Case 37 'Direção para o lado esquerdo
        
            ' Verifica se é a última posição possível para a
            'esquerda ("PosicaoX = 1") em que o bloco está
            If Posicao_na_Matriz.PosicaoX = 1 Then
            
                ' Sendo, indica-se colisão (para que o movimento do
                'bloco cesse
                DetectarColisao = True
            
            Else
            ' Não havendo colisão com "pctJogo", verifica-se
            'se hã colisão com a posição imediatamente à esquerda
            'do bloco selecionado
             
                If Posicao_na_Matriz.PosicaoX - Posicao_a_Verificar >= 1 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY, Posicao_na_Matriz.PosicaoX - Posicao_a_Verificar) <> 999 Then
                        
                        'Indica colisão
                        DetectarColisao = True
                    
                    End If
                    
                Else
                
                    'Indica colisão
                    DetectarColisao = True
                    
                End If
            
            End If

    End Select

End Function

Function Recordes(Funcao As String, Optional Posicao_do_Recorde As Integer, _
  Optional Nome_do_Jogador As String, Optional Pontuacao_do_Jogador As Integer)
'============================================================================
'  Esta Função exibe ("Funcao = Exibir") os dados armazenados em "score.lst"
'sobre os recordes dos jogadores ou armazena ("Funcao = Armazenar" e
'("Posicao_do_Recorde = <Posicao do recorde do jogador (um valor inteiro)>",
'mais o "Nome_do_Jogador" e sua "Pontuacao_do_Jogador") os dados em
'"score.lst"
' Quando "Funcao = CarregarVariavel" apenas armazenam-se os dados de "score.lst"
'em "RecordesJogo"
'============================================================================

    Dim ContadorRecordes As Integer
    Dim LinhaArquivo As String
    Dim vLinhaTXT As Long

    Select Case Funcao
    
        Case "CarregarVariavel"
        'Armazena os dados de "score.lst" em "RecordesJogo"
        
            ContadorRecordes = 1
        
            vLinhaTXT = FreeFile
    
            ' Abre o arquivo de recordes (score.lst), que está na pasta "\Data"
            Open App.Path & "\Data\score.lst" For Input As vLinhaTXT 'Abre o arquivo texto
            
            'Realiza o loop enquanto não for fim do arquivo
            Do While Not EOF(vLinhaTXT)
                
                'Lê a linha do arquivo texto onde o cursor está
                Line Input #vLinhaTXT, LinhaArquivo
           
                ' Procura pelo símbolo "%" que indica separação
                'entre o nome do jogador e sua pontuação, separando-os
                
                RecordesJogo(ContadorRecordes, 1) = Left(LinhaArquivo, (InStr(1, LinhaArquivo, "%") - 1))
                RecordesJogo(ContadorRecordes, 2) = Right(LinhaArquivo, (Len(LinhaArquivo) - InStr(1, LinhaArquivo, "%")))
                
                ContadorRecordes = ContadorRecordes + 1
            
            Loop
        
            'Fecha o Arquivo Texto
            Close vLinhaTXT
    
        Case "Exibir"
        'Exibe os valores de "score.lst" em "fraRecordes"
            
            ContadorRecordes = 0
            
            Do While ContadorRecordes <= 9
            
                frmJogo.lblJogador(ContadorRecordes) = RecordesJogo(ContadorRecordes + 1, 1)
                frmJogo.lblPontuacao(ContadorRecordes) = RecordesJogo(ContadorRecordes + 1, 2)
            
                ContadorRecordes = ContadorRecordes + 1
            
            Loop
        
        Case "Armazenar"
        ' Armazena o "Nome_do_Jogador" e sua "Pontuacao_do_Jogador"
        'em "score.lst", na "Posicao_do_Recorde" indicada
        
            ContadorRecordes = 1
        
            vLinhaTXT = FreeFile
            
            ' Acresce o novo recorde, movendo o antigo recorde da
            '"Posicao_do_Recorde" indicada UMA POSIÇÃO para baixo
            Open App.Path & "\Data\score.lst" For Output As #vLinhaTXT
            
            'Realiza o loop DEZ vezes
            Do While ContadorRecordes <= 10
            
                If ContadorRecordes = Posicao_do_Recorde Then
                
                    Print #vLinhaTXT, Nome_do_Jogador & "%" & Pontuacao_do_Jogador
                
                    'Executa um novo DO para adicionar todos os recordes
                    'que faltam abaixo deste
                    
                    ContadorRecordes = ContadorRecordes + 1
                    
                    Do While ContadorRecordes <= 10
                    
                        Print #vLinhaTXT, RecordesJogo((ContadorRecordes - 1), 1) & "%" & RecordesJogo((ContadorRecordes - 1), 2)
                    
                        ContadorRecordes = ContadorRecordes + 1
                    
                    Loop
                
                Else
                
                    Print #vLinhaTXT, RecordesJogo(ContadorRecordes, 1) & "%" & RecordesJogo(ContadorRecordes, 2)
                
                End If
            
                ContadorRecordes = ContadorRecordes + 1
                
            Loop
                
            'Fecha o Arquivo Texto
            Close vLinhaTXT
        
    End Select

End Function

Function Configuracoes(Acao As String)
'========================================================================
'  Esta Função carrega as configurações para a matriz "ConfigJogo"
'(Acao = "Carregar"); salva as configurações em "config.ini" (Acao =
'Salvar) ou salva apenas os dados de "ConfigJogo" sem aferir modificações
'(Acao = "SalvarSemAlterar").
'  Configurações existentes em "config.ini" (os números ao lado indicam
'as posições em "ConfigJogo"):
'
'   1 - SalvarConfig.Show = < True / False >
'   2 - SalvarConfig = < True / False >
'   3 - Musica = < True / False >
'   4 - Sons = < True / False >
'   5 - EstiloBlocos = < 0 (Clássico) / 1 (Novo) >
'   6 - Nivel = < 0 a 9 >
'   7 - Idioma = < Ptb / Eng >

'========================================================================

    Dim vLinhaArquivo As Long
    Dim vLinha As String
    Dim ContadorLinhaArquivo As Integer

    Select Case Acao
    
        Case "Carregar"
           
            ContadorLinhaArquivo = 1
                
            vLinhaArquivo = FreeFile
        
            ' Abre o arquivo "config.ini", que está na pasta "\Data"
            Open App.Path & "\Data\config.ini" For Input As vLinhaArquivo
            
            Do While Not EOF(vLinhaArquivo)
            
                'Lê uma linha
                Line Input #vLinhaArquivo, vLinha
            
                ' Separa o texto da linha, com base no sinal de igual "=",
                'que indica a parte com nome do parâmetro (esquerda) e o
                'seu valor (direita)
                         
                'Armazena o nome do parâmetro na primeira parte da matriz...
                ConfigJogo(ContadorLinhaArquivo, 1) = Left(vLinha, InStr(1, vLinha, "=") - 1)
                '...e seu valor na segunda parte
                ConfigJogo(ContadorLinhaArquivo, 2) = Right(vLinha, Len(vLinha) - InStr(1, vLinha, "="))

                ContadorLinhaArquivo = ContadorLinhaArquivo + 1
                  
            Loop
            
            'Fecha o Arquivo Texto
            Close vLinhaArquivo
            
            'Realiza as devidas configurações no Jogo
            
        'POSIÇÃO 3 - Músicas
            If ConfigJogo(3, 2) = True Then
                
                frmJogo.mnuMusica.Checked = True
                
            Else
            
                frmJogo.mnuMusica.Checked = False
                
            End If
            
        'POSIÇÃO 4 - Sons
            If ConfigJogo(4, 2) = True Then
                
                frmJogo.mnuSons.Checked = True
                
            Else
            
                frmJogo.mnuSons.Checked = False
                
            End If
            
        'POSIÇÃO 5 - Estilo do Bloco
            'Tira o "Checked" dos menus de Estilos
            Uncheck ("Estilo")
               
            EstiloBlocos = CInt(ConfigJogo(5, 2))
            
            Select Case EstiloBlocos
    
            Case 0
            'Muda o Estilo dos blocos para "Clássico"
            
                frmJogo.mnuEst(0).Checked = True
            
                CarregarImageList (App.Path & "\Blocos\Clássico\")
        
            Case 1
            'Muda o Estilo dos blocos para "Clássico"
                     
                frmJogo.mnuEst(1).Checked = True
                
                CarregarImageList (App.Path & "\Blocos\Novo\")
        
            End Select
            
        'POSIÇÃO 6 - Nível do Jogo
            'Tira o "Checked" dos menus de Estilos
            Uncheck ("Nível")
               
            NivelJogo = CInt(ConfigJogo(6, 2))
               
            frmJogo.mnuNivel(NivelJogo).Checked = True
            
            SelecionarNivel NivelJogo, False
            
        'POSIÇÃO 7 - Idioma
            'Tira o "Checked" dos menus de Idioma
            Uncheck ("Idioma")
        
            Select Case ConfigJogo(7, 2)
            
                Case "Ptb"
                    frmJogo.mnuIdioma(0).Checked = True
                    
                    'Armazena o Idioma do Jogo na variávle correspondente
                    IdiomaJogo = "Ptb"
                
                Case "Eng"
                    frmJogo.mnuIdioma(1).Checked = True
                    
                    'Armazena o Idioma do Jogo na variávle correspondente
                    IdiomaJogo = "Eng"
                    
            End Select
            
            TrocarIdioma (IdiomaJogo)
            
            'Coloca o novo Rótulo em "cmdNovoJogo"
            frmJogo.cmdNovoJogo.Caption = cmdNovoJogoTEXTO(cmdNovoJogoSTATUS)
            
        Case "Salvar"

            'Primeiramente, armazena os parâmetros atualmente em uso no jogo
            
            'Armazena-se se as músicas devem ser tocadas (Posição 3)
            If frmJogo.mnuMusica.Checked = True Then
            'Armazena "True" (Pode-se tocar as músicas)
                
                ConfigJogo(3, 2) = "True"
            
            Else
            'Armazena "False"
            
                ConfigJogo(3, 2) = "False"
            
            End If
            
            'Armazena-se se as sons devem ser tocados (Posição 4)
            If frmJogo.mnuSons.Checked = True Then
            'Armazena "True" (Pode-se tocar os sons)
                
                ConfigJogo(4, 2) = "True"
            
            Else
            'Armazena "False"
            
                ConfigJogo(4, 2) = "False"
            
            End If
            
            'Armazena-se o Estilo do Bloco (Posição 5)
            ConfigJogo(5, 2) = CStr(EstiloBlocos)

            'Armazena-se o Nível do Jogo (Posição 6)
            ConfigJogo(6, 2) = CStr(NivelJogo)

            'Armazena-se o Idioma do Jogo (Posição 7)
            ConfigJogo(7, 2) = IdiomaJogo
        
            vLinhaArquivo = FreeFile
        
            'Salva as configurações mo arquivo
            'Abre "config.ini" para iserção de dados
            Open App.Path & "\Data\config.ini" For Output As #vLinhaArquivo
                      
            ContadorLinhaArquivo = 1
                      
            Do While ContadorLinhaArquivo <= 7
                               
                Print #vLinhaArquivo, ConfigJogo(ContadorLinhaArquivo, 1) & "=" & ConfigJogo(ContadorLinhaArquivo, 2)
                                  
                ContadorLinhaArquivo = ContadorLinhaArquivo + 1
                                  
            Loop
                  
            'Fecha o Arquivo Texto
            Close vLinhaArquivo
        
        Case "SalvarSemAlterar"
        
            vLinhaArquivo = FreeFile
        
            'Salva as configurações mo arquivo
            'Abre "config.ini" para iserção de dados
            Open App.Path & "\Data\config.ini" For Output As #vLinhaArquivo
                      
            ContadorLinhaArquivo = 1
                      
            Do While ContadorLinhaArquivo <= 7
                               
                Print #vLinhaArquivo, ConfigJogo(ContadorLinhaArquivo, 1) & "=" & ConfigJogo(ContadorLinhaArquivo, 2)
                                  
                ContadorLinhaArquivo = ContadorLinhaArquivo + 1
                                  
            Loop
                  
            'Fecha o Arquivo Texto
            Close vLinhaArquivo
        
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
