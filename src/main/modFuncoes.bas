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
                                     
Global cmdNovoJogoSTATUS As Integer ' Indica qual a posi��o da va-
                                    'ri�vel "cmdNovoJogoTEXTO"
                                    'est� sendo usada no momento
                                    'como r�tulo em "cmdNovoJogo"
                                     
Global MsgBoxGravarRecorde As String ' Armazena o texto a ser
                                     'utilizado na MessageBox
                                     'ao indicar falha na
                                     'confec��o do nome do jogador
                                     'durante a inser��o de um
                                     'novo Recorde

Global BlocoEmJogo(179) As Boolean ' Indica se os blocos est�o
                                   'em uso em "pctJogo" ou n�o

Global EstiloBlocoEmJogo(179) As Integer ' Indica os estilos (n�mero de
                                   '1 a 7) de cada bloco existente
                                   'no jogo

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
                                   
Global RecordesJogo(10, 2) As String ' Armazena os 10 primeiros recordes
                                     'do Jogo, juntamente com o nome do
                                     'jogador
                                 
Global ConfigJogo(7, 2) As String ' Armazena as configura��es presentes
                                  '"config.ini" para posterior manipula��o

Global EstiloBlocos As Integer ' Armazena o Estilo dos blocos que est�o em
                               'Jogo
                               
Global IdiomaJogo As String 'Armazena o Idioma do Jogo

Global NivelJogo As Integer 'Armazena o N�vel do Jogo

Global MusicaExecutando As Boolean ' Indica se M�sicas MIDI est�o
                                   'sendo executadas ou n�o
                        
Global TempoMusica As Integer 'Armazena o Tempo Inicial da M�sica em execu��o
                        
                        
                        
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
                 
Public Enum SndPlayFlags ' Enumera��o utilizada na fun��o
  SND_SYNC = &H0         'para tocar arquivos WAVE
  SND_ASYNC = &H1
  SND_NODEFAULT = &H2
  SND_MEMORY = &H4
  SND_LOOP = &H8
  SND_NOSTOP = &H10
End Enum

Public Enum AcaoTocador ' Enumera��o utilizada na fun��o
  Tocar                 'para tocar arquivos MIDI
  Pausar
  Resumir
  Parar
  Tempo
End Enum



'Fun��o para tocar arquivos WAVE
Public Declare Function sndPlaySound Lib "winmm.dll" _
       Alias "sndPlaySoundA" (ByVal lpszSoundName As _
       String, ByVal uFlags As Long) As Long
       
'Fun��o para tocar arquivos MIDI
Public Declare Function mciExecute Lib "winmm.dll" _
       (ByVal lpstrCommand As String) As Long
       
'Fun��o para uso da convers�o de caminhos longos de arquivos
'Extra�da de FreeVBCode.com (http://www.freevbcode.com)
Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
    As Long
   
   
   
Function RetornarCaminhoDOS(ByVal Caminho As String) As String
'==============================================================
' Esta fun��o retorna um caminho de arquivo no formato 8.3
'(DOS) dado este no formato atual de 255 caracteres, retornando
'"" caso o caminho n�o exista.
'
' Extra�da de FreeVBCode.com (http://www.freevbcode.com)
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
' Esta fun��o toca um arquivo WAVE
'
' Op��es do par�metro "Flags":
' - "SND_SYNC" - Toca o WAVE sincronizado (default).
' - "SND_ASYNC" - Toca o WAVE sem sincronismo.
' - "SND_NODEFAULT" - N�o usa o som padr�o.
' - "SND_MEMORY" - Direciona o "IpszSoundName" para o lugar
'    de um arquivo na mem�ria.
' - "SND_LOOP" - Toca o WAVE em looping, at� que se o mande
'    parar
' - "SND_NOSTOP" - N�o p�ra o som que estiver tocando.
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
' Esta fun��o coordena o tocador de MIDI, utilizando-se para
'tal os par�metros presentes na fun��o.
' Par�metros:
'   - "Acao": Baseado na enumera��o "AcaoTocador", possui as
'     seguintes op��es:
'          * "Tocar" - Abre uma m�sica MIDI, usando como base
'            o par�metro "Arquivo
'          * "Pausar" - Pausa uma m�sica MIDI em execu��o
'          * "Resumir" - Reinicia uma m�sica MIDI anteriormente
'            pausada
'          * "Parar" - Fecha o arquivo MIDI
'          * "Tempo" - aumenta ou diminui o Tempo da
'            m�sica MIDIcom base no par�metro "Tempo"
'
' Obs.: "MusicaMIDI" � o nome dado ao processo que aciona o
'arquivo MIDI
'===============================================================
     
    Select Case Acao
    
        Case 0
            mciExecute "open " & RetornarCaminhoDOS(ArquivoMIDI) & " type sequencer alias Musica"
            
            mciExecute "play Musica"
            
            'Indica que uma m�sica est� sendo tocada
            MusicaExecutando = True
        
        Case 1
            mciExecute "pause Musica"
        
        Case 2
            mciExecute "play Musica"
        
        Case 3
            'pergunta se h� uma m�sica em execu��o para fech�-la
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
    Dim Objeto_a_Manipular As Object
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
                    Set Objeto_a_Manipular = frmJogo(Left(vControle, Len(vControle) - 1))
                    Objeto_a_Manipular(CInt(Right(vControle, 1))).Caption = vLabel
        
                Else
                
                    'Modifica o r�tulo do corrente objeto o idioma selecionado
                    Set Objeto_a_Manipular = frmJogo(vControle)
                    Objeto_a_Manipular.Caption = vLabel
        
                End If
         
            Case "*frmSobre*"
    
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o r�tulo do controle
                '*Usa como base para separa��o da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))

                'Verifica se o controle atual � o pr�prio Formul�rio
                If vControle = "frmSobre" Then
                
                    'Modifica o r�tulo do Formul�rio
                    frmSobre.Caption = vLabel
                        
                Else
                    
                    'Modifica o r�tulo do corrente objeto o idioma selecionado
                    Set Objeto_a_Manipular = frmSobre(vControle)
                    Objeto_a_Manipular.Caption = vLabel
        
                End If
                
            Case "*frmSalvarConfig*"
    
                'Desmembra a linha em:
                '   - vControle: nome do controle
                '   - vLabel: o r�tulo do controle
                '*Usa como base para separa��o da linha o sinal de
                'igual(=)
                vControle = Left(Linha, InStr(1, Linha, "=") - 1)
                vLabel = Right(Linha, Len(Linha) - InStr(1, Linha, "="))

                'Modifica o r�tulo do corrente objeto o idioma selecionado
                Set Objeto_a_Manipular = frmSalvarConfig(vControle)
                Objeto_a_Manipular.Caption = vLabel
                       
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
           
                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco
                    
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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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

                    'Armazena o Estilo do Bloco em "EstiloBlocoEmJogo"
                    EstiloBlocoEmJogo(IndexBloco) = EstiloBloco

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
    
    Dim ContadorProxIndex As Integer
    
    Select Case ContainerDoBloco
    
        Case "pctJogo"
    
            ContadorProxIndex = 0
            
            Do While ContadorProxIndex < 180
            
                If BlocoEmJogo(ContadorProxIndex) = False Then
                
                    ProxIndexBloco = ContadorProxIndex
                    
                    'Indica que este bloco est� em uso no Jogo
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
    
    Do While Contador_da_Reordem >= 1
    
        Contador_das_Colunas = 1
    
        Do While Contador_das_Colunas <= 10
        
            'Verifica se h� bloco nesta posi��o
            If Jogo(Contador_da_Reordem, Contador_das_Colunas) <> 999 Then
            
                'Move o bloco uma posi��o abaixo
                frmJogo.imgBloco(Jogo(Contador_da_Reordem, Contador_das_Colunas)).Top = (Contador_da_Reordem) * UmaPosicao
                'Indica a nova posi��o do bloco na matrtiz
                Jogo(Contador_da_Reordem + 1, Contador_das_Colunas) = Jogo(Contador_da_Reordem, Contador_das_Colunas)
                'Armazena "999" na antiga posi��o do bloco na matriz
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
'  Esta Fun��o move os blocos de �ndices "I1", "I2", "I3" e "I4" para:
'       - Cima ("Direcao = 38")
'       - Baixo ("Direcao = 40")
'       - Esquerda ("Direcao = 37")
'       - Direita ("Direcao = 39")
'  Como a Fun��o come�a sempre o movimento do Bloco a partir da posi��o 1
'de  "Blocos_a_mover()", h� a verifica��o de quais posi��es devem ser
'movidas primeiro (atrav�s de "TipoBloco_a_Mover" e "PosicaoDoBloco_a_Mover")
'para evitar sobreposi��o de  blocos.
'============================================================================

    Dim Blocos_a_mover(4) As Integer
    Dim Contador_MoverBloco As Integer
    Dim Posicao_Bloco_Matriz As Posicao
    'Armazenam as varia��es de posi��es para a manipula��o na matriz "Jogo"
    Dim PosicaoMatriz_X As Integer, PosicaoMatriz_Y As Integer
    'Armazenam as varia��es de posi��es para a manipula��o dos blocos
    Dim PosicaoBloco_X As Integer, PosicaoBloco_Y As Integer

    Select Case Direcao
    
        Case 38
        'Move os blocos uma posi��o para cima
     
            PosicaoMatriz_X = 0
            PosicaoMatriz_Y = -1
            PosicaoBloco_X = 0
            PosicaoBloco_Y = -UmaPosicao
                        
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
                
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
        'Move os blocos uma posi��o para baixo
     
            PosicaoMatriz_X = 0
            PosicaoMatriz_Y = 1
            PosicaoBloco_X = 0
            PosicaoBloco_Y = UmaPosicao
            
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
                
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
        'Move os blocos uma posi��o para a esquerda
        
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
                            
                    'Verifica a posi��o atual do bloco
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
                
                    'Verifica a posi��o atual do bloco
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
        'Move os blocos uma posi��o para a direita
            
            PosicaoMatriz_X = 1
            PosicaoMatriz_Y = 0
            PosicaoBloco_X = UmaPosicao
            PosicaoBloco_Y = 0
            
            Select Case TipoBloco_a_Mover
            
                Case 0
                'Blocos do tipo ****
                            
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
                            
                    'Verifica a posi��o atual do bloco
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
                
                    'Verifica a posi��o atual do bloco
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
    
            'Move o bloco uma posi��o acima/abaixo/esquerda/direita na matriz
            Posicao_Bloco_Matriz = LocalizarBlocoNaMatriz(Blocos_a_mover(Contador_MoverBloco))
            Jogo(Posicao_Bloco_Matriz.PosicaoY + PosicaoMatriz_Y, Posicao_Bloco_Matriz.PosicaoX + PosicaoMatriz_X) = Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX)
            Jogo(Posicao_Bloco_Matriz.PosicaoY, Posicao_Bloco_Matriz.PosicaoX) = 999
                    
            'Move o bloco 1 posi��o acima/abaixo/esquerda/direita em "pctJogo"
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Top + PosicaoBloco_Y)
            frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left = (frmJogo.imgBloco(Blocos_a_mover(Contador_MoverBloco)).Left + PosicaoBloco_X)
                
        End If
            
        Contador_MoverBloco = Contador_MoverBloco + 1
            
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
        
            'Verifica qual o tipo de Bloco
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
                            ArmazenarPosicoes.idx3 = 4
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
                            'abaixo dos blocos "2", "3" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posi��es diretamente
                            'abaixo dos blocos "1" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 1
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
        
            'Verifica qual o tipo de Bloco
            Select Case Tipo_BlocoEmJogo
            
                Case 0
                    'Blocos do tipo ****
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posi��o "1"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posi��es do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 4
                        
                    End Select
                    
                Case 1
                    'Blocos do tipo ***
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posi��o "1"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posi��es do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 2
                    '                *
                    'Blocos do tipo ***
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "1", "2" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 3
                    '                 *
                    'Blocos do tipo ***
        
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "1", "2" e "3"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posi��es diretamente
                            '� esquerda dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 4
                    '                **
                    'Blocos do tipo **
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica apenas as posi��es diretamente
                            '� esquerda dos blocos "1" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posi��es diretamente
                            '� esquerda dos blocos "1", "2" e "3" do conjunto
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 5
                    '               **
                    'Blocos do tipo **
                    
                    ' Neste tipo de bloco, apenas as posi��es "1" e "3"
                    's�o sempre verificadas
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
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posi��o "4"
                            ArmazenarPosicoes.idx1 = 4
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posi��es do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 4
                        
                    End Select
                    
                Case 1
                    'Blocos do tipo ***
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            'Verifica apenas a posi��o "3"
                            ArmazenarPosicoes.idx1 = 3
                            ArmazenarPosicoes.idx2 = 999
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica todas as posi��es do conjunto
                            'de blocos
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 2
                    '                *
                    'Blocos do tipo ***
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "1", "2" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 3
                    '                 *
                    'Blocos do tipo ***
        
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "1" e "4"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "2", "3" e "4"
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                            
                        Case 2
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "1" e "2"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 3
                            ' Verifica as posi��es diretamente
                            '� direita dos blocos "1", "2" e "3"
                            ArmazenarPosicoes.idx1 = 1
                            ArmazenarPosicoes.idx2 = 2
                            ArmazenarPosicoes.idx3 = 3
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                
                Case 4
                    '                **
                    'Blocos do tipo **
                    
                    'Verifica a posi��o atual do bloco
                    Select Case PosicaoAtual_BlocoEmJogo
                    
                        Case 0
                            ' Verifica apenas as posi��es diretamente
                            '� direita dos blocos "2" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 4
                            ArmazenarPosicoes.idx3 = 999
                            ArmazenarPosicoes.idx4 = 999
                        
                        Case 1
                            ' Verifica apenas as posi��es diretamente
                            '� direita dos blocos "1", "3" e "4" do conjunto
                            ArmazenarPosicoes.idx1 = 2
                            ArmazenarPosicoes.idx2 = 3
                            ArmazenarPosicoes.idx3 = 4
                            ArmazenarPosicoes.idx4 = 999
                        
                    End Select
                    
                Case 5
                    '               **
                    'Blocos do tipo **
                    
                    ' Neste tipo de bloco, apenas as posi��es "2" e "4"
                    's�o sempre verificadas
                    ArmazenarPosicoes.idx1 = 2
                    ArmazenarPosicoes.idx2 = 4
                    ArmazenarPosicoes.idx3 = 999
                    ArmazenarPosicoes.idx4 = 999
                    
            End Select
        
    End Select

End Function

Function MenuNivel(Acao As String)
'=======================================================================
'  Esta Fun��o habilita (Acao = "Habilitar") os menus referentes �
'sele��o de N�vel de Jogo ou desabilita-os (Acao = "Desabilitar")
'=======================================================================

    Dim ContadorNivel As Integer
    
    ContadorNivel = 0

    Select Case Acao
    
        Case "Habilitar"

            'Habilita os menus selecionadores de N�vel de Jogo
            Do While ContadorNivel <= 9
                
                frmJogo.mnuNivel(ContadorNivel).Enabled = True
                
                ContadorNivel = ContadorNivel + 1
                
            Loop
    
        Case "Desabilitar"
        
            'Desabilita os menus selecionadores de N�vel de Jogo
            Do While ContadorNivel <= 9
                
                frmJogo.mnuNivel(ContadorNivel).Enabled = False
                
                ContadorNivel = ContadorNivel + 1
                
            Loop
        
        
    End Select

End Function

Function SelecionarNivel(Nivel_do_Jogo As Integer, EmJogo As Boolean)
'=======================================================================
'  Esta Fun��o seleciona o n�vel do jogo com base no n�mero informado
'atrav�s de "Nivel_do_Jogo" alterando, inclusive, as indica��es de
'n�vel selecionado no menu e em "lblPontos".
'Obs.: a informa��o de n�vel nos menus s� ser� alterada se
'"EmJogo" = "False" (indicando que a  mudan�a de n�vel ocorre por
'interven��o do usu�rio e n�o da Engine do Jogo
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
'  Esta Fun��o detecta a colis�o, em "pctJogo", do bloco de "Indice_Bloco"
'selecionado analisando a "Direcao_a_Verificar" escolhida
'  O par�metro "Posicao_a_Verificar" indica quantas posi��es abaixo do
'bloco de "Indice_Bloco" selecionado devem ser puladas para
'a verifica��o. O padr�o � "1" (indica que a posi��o imediatamente
'acima/abaixo/� esquerda/� direita do bloco deve ser verificada;
'caso fosse "2", por exemplo, a fun��o desconsideraria a posi��o
'imediatamente acima/abaixo/� esquerda/� direita e analisaria a
'pr�xima posi��o ap�s esta)
'=========================================================================

    Dim Posicao_na_Matriz As Posicao

    'Caso "Posicao_a_Verificar" seja igual a "0" (par�metro
    'n�o informado), esta passa a valer "1" (valor padr�o)
    If Posicao_a_Verificar = 0 Then
    
        Posicao_a_Verificar = 1
        
    End If

    'A princ�pio, n�o h� colis�o do bloco
    DetectarColisao = False

    'Detecta a posi��o do bloco de "Indice_Bloco" selecionado
    Posicao_na_Matriz = LocalizarBlocoNaMatriz(Indice_Bloco)

    'Detecta a colis�o na dire��o selecionada
    Select Case Direcao_a_Verificar
    
        Case 38 'Dire��o para cima
            
            ' Verifica se � a primeira linha ("PosicaoY = 1")
            'em que o bloco est�
            If Posicao_na_Matriz.PosicaoY = 1 Then
        
                ' Sendo, indica-se colis�o (uma vez que n�o h� como
                'o Bloco subir
                DetectarColisao = True
            
            Else
            ' N�o havendo colis�o com "pctJogo", verifica-se
            'se h� colis�o com a posi��o imediatamente acima
            'do bloco selecionado
            
                If Posicao_na_Matriz.PosicaoY - Posicao_a_Verificar >= 1 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY - Posicao_a_Verificar, Posicao_na_Matriz.PosicaoX) <> 999 Then
                        
                        'Indica colis�o
                        DetectarColisao = True
                    
                    End If
                
                Else
                
                    'Indica colis�o
                    DetectarColisao = True
                
                End If
            
            End If
    
        Case 40 'Dire��o para baixo
            
            ' Verifica se � a �ltima linha ("PosicaoY = 18")
            'em que o bloco est�
            If Posicao_na_Matriz.PosicaoY = 18 Then
        
                ' Sendo, indica-se colis�o (para que o movimento do
                'bloco cesse
                DetectarColisao = True
            
            Else
            ' N�o havendo colis�o com "pctJogo", verifica-se
            'se h� colis�o com a posi��o imediatamente abaixo
            'do bloco selecionado
            
                If Posicao_na_Matriz.PosicaoY + Posicao_a_Verificar <= 18 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY + Posicao_a_Verificar, Posicao_na_Matriz.PosicaoX) <> 999 Then
                        
                        'Indica colis�o
                        DetectarColisao = True
                    
                    End If
                
                Else
                
                    'Indica colis�o
                    DetectarColisao = True
                
                End If
            
            End If
    
        Case 39 'Dire��o para o lado direito
            
            ' Verifica se � a �ltima posi��o poss�vel para a
            'direita ("PosicaoX = 10") em que o bloco est�
            If Posicao_na_Matriz.PosicaoX = 10 Then
            
                ' Sendo, indica-se colis�o (para que o movimento do
                'bloco cesse
                DetectarColisao = True
            
            Else
            ' N�o havendo colis�o com "pctJogo", verifica-se
            'se h� colis�o com a posi��o imediatamente � direita
            'do bloco selecionado
                
                If Posicao_na_Matriz.PosicaoX + Posicao_a_Verificar <= 10 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY, Posicao_na_Matriz.PosicaoX + Posicao_a_Verificar) <> 999 Then
                        
                        'Indica colis�o
                        DetectarColisao = True
                    
                    End If
                
                Else
                
                    'Indica colis�o
                    DetectarColisao = True
                
                End If
            
            End If

        Case 37 'Dire��o para o lado esquerdo
        
            ' Verifica se � a �ltima posi��o poss�vel para a
            'esquerda ("PosicaoX = 1") em que o bloco est�
            If Posicao_na_Matriz.PosicaoX = 1 Then
            
                ' Sendo, indica-se colis�o (para que o movimento do
                'bloco cesse
                DetectarColisao = True
            
            Else
            ' N�o havendo colis�o com "pctJogo", verifica-se
            'se h� colis�o com a posi��o imediatamente � esquerda
            'do bloco selecionado
             
                If Posicao_na_Matriz.PosicaoX - Posicao_a_Verificar >= 1 Then
            
                    If Jogo(Posicao_na_Matriz.PosicaoY, Posicao_na_Matriz.PosicaoX - Posicao_a_Verificar) <> 999 Then
                        
                        'Indica colis�o
                        DetectarColisao = True
                    
                    End If
                    
                Else
                
                    'Indica colis�o
                    DetectarColisao = True
                    
                End If
            
            End If

    End Select

End Function

Function Recordes(Funcao As String, Optional Posicao_do_Recorde As Integer, _
  Optional Nome_do_Jogador As String, Optional Pontuacao_do_Jogador As Integer)
'============================================================================
'  Esta Fun��o exibe ("Funcao = Exibir") os dados armazenados em "score.lst"
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
    
            ' Abre o arquivo de recordes (score.lst), que est� na pasta "\Data"
            Open App.Path & "\Data\score.lst" For Input As vLinhaTXT 'Abre o arquivo texto
            
            'Realiza o loop enquanto n�o for fim do arquivo
            Do While Not EOF(vLinhaTXT)
                
                'L� a linha do arquivo texto onde o cursor est�
                Line Input #vLinhaTXT, LinhaArquivo
           
                ' Procura pelo s�mbolo "%" que indica separa��o
                'entre o nome do jogador e sua pontua��o, separando-os
                
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
            '"Posicao_do_Recorde" indicada UMA POSI��O para baixo
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
'  Esta Fun��o carrega as configura��es para a matriz "ConfigJogo"
'(Acao = "Carregar"); salva as configura��es em "config.ini" (Acao =
'Salvar) ou salva apenas os dados de "ConfigJogo" sem aferir modifica��es
'(Acao = "SalvarSemAlterar").
'  Configura��es existentes em "config.ini" (os n�meros ao lado indicam
'as posi��es em "ConfigJogo"):
'
'   1 - SalvarConfig.Show = < True / False >
'   2 - SalvarConfig = < True / False >
'   3 - Musica = < True / False >
'   4 - Sons = < True / False >
'   5 - EstiloBlocos = < 0 (Cl�ssico) / 1 (Novo) >
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
        
            ' Abre o arquivo "config.ini", que est� na pasta "\Data"
            Open App.Path & "\Data\config.ini" For Input As vLinhaArquivo
            
            Do While Not EOF(vLinhaArquivo)
            
                'L� uma linha
                Line Input #vLinhaArquivo, vLinha
            
                ' Separa o texto da linha, com base no sinal de igual "=",
                'que indica a parte com nome do par�metro (esquerda) e o
                'seu valor (direita)
                         
                'Armazena o nome do par�metro na primeira parte da matriz...
                ConfigJogo(ContadorLinhaArquivo, 1) = Left(vLinha, InStr(1, vLinha, "=") - 1)
                '...e seu valor na segunda parte
                ConfigJogo(ContadorLinhaArquivo, 2) = Right(vLinha, Len(vLinha) - InStr(1, vLinha, "="))

                ContadorLinhaArquivo = ContadorLinhaArquivo + 1
                  
            Loop
            
            'Fecha o Arquivo Texto
            Close vLinhaArquivo
            
            'Realiza as devidas configura��es no Jogo
            
        'POSI��O 3 - M�sicas
            If ConfigJogo(3, 2) = True Then
                
                frmJogo.mnuMusica.Checked = True
                
            Else
            
                frmJogo.mnuMusica.Checked = False
                
            End If
            
        'POSI��O 4 - Sons
            If ConfigJogo(4, 2) = True Then
                
                frmJogo.mnuSons.Checked = True
                
            Else
            
                frmJogo.mnuSons.Checked = False
                
            End If
            
        'POSI��O 5 - Estilo do Bloco
            'Tira o "Checked" dos menus de Estilos
            Uncheck ("Estilo")
               
            EstiloBlocos = CInt(ConfigJogo(5, 2))
            
            Select Case EstiloBlocos
    
            Case 0
            'Muda o Estilo dos blocos para "Cl�ssico"
            
                frmJogo.mnuEst(0).Checked = True
            
                CarregarImageList (App.Path & "\Blocos\Cl�ssico\")
        
            Case 1
            'Muda o Estilo dos blocos para "Cl�ssico"
                     
                frmJogo.mnuEst(1).Checked = True
                
                CarregarImageList (App.Path & "\Blocos\Novo\")
        
            End Select
            
        'POSI��O 6 - N�vel do Jogo
            'Tira o "Checked" dos menus de Estilos
            Uncheck ("N�vel")
               
            NivelJogo = CInt(ConfigJogo(6, 2))
               
            frmJogo.mnuNivel(NivelJogo).Checked = True
            
            SelecionarNivel NivelJogo, False
            
        'POSI��O 7 - Idioma
            'Tira o "Checked" dos menus de Idioma
            Uncheck ("Idioma")
        
            Select Case ConfigJogo(7, 2)
            
                Case "Ptb"
                    frmJogo.mnuIdioma(0).Checked = True
                    
                    'Armazena o Idioma do Jogo na vari�vle correspondente
                    IdiomaJogo = "Ptb"
                
                Case "Eng"
                    frmJogo.mnuIdioma(1).Checked = True
                    
                    'Armazena o Idioma do Jogo na vari�vle correspondente
                    IdiomaJogo = "Eng"
                    
            End Select
            
            TrocarIdioma (IdiomaJogo)
            
            'Coloca o novo R�tulo em "cmdNovoJogo"
            frmJogo.cmdNovoJogo.Caption = cmdNovoJogoTEXTO(cmdNovoJogoSTATUS)
            
        Case "Salvar"

            'Primeiramente, armazena os par�metros atualmente em uso no jogo
            
            'Armazena-se se as m�sicas devem ser tocadas (Posi��o 3)
            If frmJogo.mnuMusica.Checked = True Then
            'Armazena "True" (Pode-se tocar as m�sicas)
                
                ConfigJogo(3, 2) = "True"
            
            Else
            'Armazena "False"
            
                ConfigJogo(3, 2) = "False"
            
            End If
            
            'Armazena-se se as sons devem ser tocados (Posi��o 4)
            If frmJogo.mnuSons.Checked = True Then
            'Armazena "True" (Pode-se tocar os sons)
                
                ConfigJogo(4, 2) = "True"
            
            Else
            'Armazena "False"
            
                ConfigJogo(4, 2) = "False"
            
            End If
            
            'Armazena-se o Estilo do Bloco (Posi��o 5)
            ConfigJogo(5, 2) = CStr(EstiloBlocos)

            'Armazena-se o N�vel do Jogo (Posi��o 6)
            ConfigJogo(6, 2) = CStr(NivelJogo)

            'Armazena-se o Idioma do Jogo (Posi��o 7)
            ConfigJogo(7, 2) = IdiomaJogo
        
            vLinhaArquivo = FreeFile
        
            'Salva as configura��es mo arquivo
            'Abre "config.ini" para iser��o de dados
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
        
            'Salva as configura��es mo arquivo
            'Abre "config.ini" para iser��o de dados
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
