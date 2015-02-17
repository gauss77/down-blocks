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
Global NumBloco As Long 'Indica o número do índice do Bloco que será
                        'utilizado no Jogo ("imgBloco(índice)")

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

'On Error GoTo Erro

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
        Line Input #vLinhaTXT, linha
   
        ' Verifica se o texto da corrente linha do arquivo é
        'indicação de nome de formulário
        If Left(linha, 1) = "*" Then
        
            Formulario = linha
            ' Lê a próxima linha do arquivo texto (que indica
            'um controle do corrente formulário)
            Line Input #vLinhaTXT, linha
            
        End If
   
        'Verifica de qual formulário pertence o atual componente
        Select Case Formulario
        
            Case "*Variaveis*"
                vControle = Left(linha, InStr(1, linha, "=") - 1)
                vLabel = Right(linha, Len(linha) - InStr(1, linha, "="))
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
                vControle = Left(linha, InStr(1, linha, "=") - 1)
                vLabel = Right(linha, Len(linha) - InStr(1, linha, "="))
                
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
                vControle = Left(linha, InStr(1, linha, "=") - 1)
                vLabel = Right(linha, Len(linha) - InStr(1, linha, "="))

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

    ' Primeiramente, seleciona o tipo do bloco que será formado,
    'através da função "Random", escolhendo um
    'número que esteja entre "0" e "4"
    
    'Indica qual PictureBox utilizar, através do Container
    If Container = "pctJogo" Then
    
        'Indica a qual PictureBox se referenciará "pctObj"
        Set pctObj = frmJogo.imgBloco
        IndexBloco = NumBloco
    
    Else
        
        'Indica a qual PictureBox se referenciará "pctObj"
        'Set pctObj = frmJogo.pctProxBloc
        Set pctObj = frmJogo.imgBlocoProx
        
        ' "IndexBloco" valerá "0" por conta de os Blocos em
        '"pctProx" serem sempre os mesmos a serem utilizados,
        'partindo-se sempre do primeiro Bloco (índice "0")
        IndexBloco = 0
    
        'Esconde todos os Blocos utilizados em "pctProx"
        pctObj(0).Visible = False
        pctObj(1).Visible = False
        pctObj(2).Visible = False
        pctObj(3).Visible = False
    
    End If

    ' Assegura que o valor de EstiloBloco esteja sempre entre
    '"1" e "7"
    If EstiloBloco = 0 Then EstiloBloco = 1

    Select Case TipoBloco
    
        Case 0
            'Monta um conjunto de Blocos do tipo ****, no Container escolhido
            Contador = 3
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições iniciais do primeiro bloco
                PosX = 1125
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro bloco
                PosX = 150
                PosY = 237
            
            End If

            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375

                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
                    
        Case 1
            'Monta um conjunto de Blocos do tipo ***, no Container escolhido
            Contador = 2
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições iniciais do primeiro bloco
                PosX = 1125
                PosY = 0
            
            Else
            
                'Define as posições iniciais do primeiro Bloco
                PosX = 337
                PosY = 237
                
            End If

            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375

                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
        
        Case 2
            '                                     *
            'Monta um conjunto de Blocos do tipo ***, no Container escolhido
           
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                ' Define as posições iniciais do primeiro Bloco (neste
                'caso, do bloco que fica acima dos outros três)
                PosX = 1500
                PosY = 0
            
            Else
           
                ' Define as posições iniciais do primeiro Bloco (neste
                'caso, do bloco que fica acima dos outros três)
                PosX = 712
                PosY = 50
                
            End If
            
            'Insere o primeiro Bloco
            pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
            pctObj(IndexBloco).Top = PosY
            pctObj(IndexBloco).Left = PosX
            pctObj(IndexBloco).Visible = True
            
            IndexBloco = IndexBloco + 1
            
            'Insere os outros três Blocos
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                ' Define as posições
                PosX = 1125
                PosY = 375
            
            Else
            
                'Define as posições
                PosX = 337
                PosY = 425
                
            End If

            'Corrige os valores do contador
            Contador = 2
            
            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375
                
                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
        
        Case 3
            '                                      *
            'Monta um conjunto de Blocos do tipo ***, no Container escolhido
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                ' Define as posições iniciais do primeiro Bloco (neste
                'caso, do bloco que fica acima dos outros três)
                PosX = 1875
                PosY = 0
            
            Else
            
                ' Define as posições iniciais do primeiro Bloco (neste
                'caso, do bloco que fica acima dos outros três)
                PosX = 1087
                PosY = 50
                
            End If
            
            'Insere o primeiro Bloco
            pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
            pctObj(IndexBloco).Top = PosY
            pctObj(IndexBloco).Left = PosX
            pctObj(IndexBloco).Visible = True
            
            IndexBloco = IndexBloco + 1
            
            'Insere os outros três Blocos
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                ' Define as posições
                PosX = 1125
                PosY = 375
            
            Else
            
                'Define as posições
                PosX = 337
                PosY = 425
                
            End If

            'Corrige os valores do contador
            Contador = 2
            
            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375
    
                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
            
        Case 4
            '                                     **
            'Monta um conjunto de Blocos do tipo ** , no Container escolhido
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições iniciais dos dois primeiros Blocos
                PosX = 1500
                PosY = 0
            
            Else
            
                'Define as posições iniciais dos dois primeiros Blocos
                PosX = 712
                PosY = 50
                
            End If

            'Corrige os valores do contador
            Contador = 1

            'Executa o primeiro DO
            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375
                
                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Redefine as posições iniciais dos dois primeiros Blocos
                PosX = 1125
                PosY = 375
            
            Else
            
                'Redefine as posições iniciais dos dois primeiros Blocos
                PosX = 337
                PosY = 425
                
            End If
            
            'Corrige os valores do contador
            Contador = 1
            
            'Executa o segundo DO
            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375
                
                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
            
        Case 5
            '                                    **
            'Monta um conjunto de Blocos do tipo **, no Container escolhido
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Define as posições iniciais dos dois primeiros Blocos
                PosX = 1500
                PosY = 0
            
            Else
            
                'Define as posições iniciais dos dois primeiros Blocos
                PosX = 525
                PosY = 50
                
            End If

            'Corrige os valores do contador
            Contador = 1

            'Executa o primeiro DO
            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375
                
                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
            
            'Verifica qual o Container dos Blocos
            If Container = "pctJogo" Then
                
                'Redefine as posições iniciais dos dois primeiros Blocos
                PosX = 1500
                PosY = 375
            
            Else
            
                'Redefine as posições iniciais dos dois primeiros Blocos
                PosX = 525
                PosY = 425
                
            End If
            
            'Corrige os valores do contador
            Contador = 1
            
            'Executa o segundo DO
            Do While Contador >= 0
           
                pctObj(IndexBloco).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                pctObj(IndexBloco).Top = PosY
                pctObj(IndexBloco).Left = PosX
                pctObj(IndexBloco).Visible = True
                
                PosX = PosX + 375
                
                IndexBloco = IndexBloco + 1
                Contador = Contador - 1
                
            Loop
            
    End Select
    
    If Container = "pctJogo" Then
    
        ' Se o Container pedido for "pctJogo", indica o próximo
        'valor que se referenciará aos Blocos em "NumBloco"
        NumBloco = IndexBloco
    
    End If
    
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

