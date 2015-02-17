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
                        'utilizado no Jogo ("pctBloco(índice)")

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

Function GerarBlocos()
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
'     **
'   - **; representado pelo número "4"
'=============================================================

    Dim Contador As Integer
    Dim PosX As Integer, PosY As Integer
    Dim EstiloBloco As Integer

    ' Primeiramente, seleciona o tipo do bloco que será formado,
    'através da função "Random", escolhendo um
    'número que esteja entre "0" e "4"
    
    'Esconde todos os Blocos utilizados em "pctProx"
    frmJogo.pctProxBloc(0).Visible = False
    frmJogo.pctProxBloc(1).Visible = False
    frmJogo.pctProxBloc(2).Visible = False
    frmJogo.pctProxBloc(3).Visible = False

    Select Case Random(4)
    
        Case 0
            'Monta um conjunto de Blocos do tipo ****, em "pctProx"
            Contador = 3
            
            'Define as posições iniciais do primeiro bloco
            PosX = 150
            PosY = 237
            
            ' Seleciona um número que se referenciará a um
            'estilo do Bloco (imagens armazenadas com valores
            'de "1" a "7" em "ImageListBlocos")
            EstiloBloco = Random(7)
            If EstiloBloco = 0 Then EstiloBloco = 1

            Do While Contador >= 0
           
                frmJogo.pctProxBloc(Contador).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                frmJogo.pctProxBloc(Contador).Top = PosY
                frmJogo.pctProxBloc(Contador).Left = PosX
                frmJogo.pctProxBloc(Contador).Visible = True
                
                PosX = PosX + 375

                Contador = Contador - 1
                
            Loop
                    
        Case 1
            'Monta um conjunto de Blocos do tipo ***, em "pctProx"
            Contador = 2
            
            'Define as posições iniciais do primeiro Bloco
            PosX = 337
            PosY = 237
            
            ' Seleciona um número que se referenciará a um
            'estilo do Bloco (imagens armazenadas com valores
            'de "1" a "7" em "ImageListBlocos")
            EstiloBloco = Random(7)
            If EstiloBloco = 0 Then EstiloBloco = 1

            Do While Contador >= 0
           
                frmJogo.pctProxBloc(Contador).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                frmJogo.pctProxBloc(Contador).Top = PosY
                frmJogo.pctProxBloc(Contador).Left = PosX
                frmJogo.pctProxBloc(Contador).Visible = True
                
                PosX = PosX + 375

                Contador = Contador - 1
                
            Loop
        
        Case 2
            '                                     *
            'Monta um conjunto de Blocos do tipo ***, em "pctProx"
           
            ' Define as posições iniciais do primeiro Bloco (neste
            'caso, do bloco que fica acima dos outros três)
            PosX = 712
            PosY = 50
            
            ' Seleciona um número que se referenciará a um
            'estilo do Bloco (imagens armazenadas com valores
            'de "1" a "7" em "ImageListBlocos")
            EstiloBloco = Random(7)
            If EstiloBloco = 0 Then EstiloBloco = 1
            
            'Insere o primeiro Bloco
            frmJogo.pctProxBloc(3).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
            frmJogo.pctProxBloc(3).Top = PosY
            frmJogo.pctProxBloc(3).Left = PosX
            frmJogo.pctProxBloc(3).Visible = True
            
            'Insere os outros três Blocos
            
            'Define as posições
            PosX = 337
            PosY = 425

            'Corrige os valores do contador
            Contador = 2
            
            Do While Contador >= 0
           
                frmJogo.pctProxBloc(Contador).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                frmJogo.pctProxBloc(Contador).Top = PosY
                frmJogo.pctProxBloc(Contador).Left = PosX
                frmJogo.pctProxBloc(Contador).Visible = True
                
                PosX = PosX + 375
                
                Contador = Contador - 1
                
            Loop
        
        Case 3
            '                                      *
            'Monta um conjunto de Blocos do tipo ***, em "pctProx"
            
            ' Define as posições iniciais do primeiro Bloco (neste
            'caso, do bloco que fica acima dos outros três)
            PosX = 1087
            PosY = 50
            
            ' Seleciona um número que se referenciará a um
            'estilo do Bloco (imagens armazenadas com valores
            'de "1" a "7" em "ImageListBlocos")
            EstiloBloco = Random(7)
            If EstiloBloco = 0 Then EstiloBloco = 1
            
            'Insere o primeiro Bloco
            frmJogo.pctProxBloc(3).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
            frmJogo.pctProxBloc(3).Top = PosY
            frmJogo.pctProxBloc(3).Left = PosX
            frmJogo.pctProxBloc(3).Visible = True
            
            'Insere os outros três Blocos
            
            'Define as posições
            PosX = 337
            PosY = 425

            'Corrige os valores do contador
            Contador = 2
            
            Do While Contador >= 0
           
                frmJogo.pctProxBloc(Contador).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                frmJogo.pctProxBloc(Contador).Top = PosY
                frmJogo.pctProxBloc(Contador).Left = PosX
                frmJogo.pctProxBloc(Contador).Visible = True
                
                PosX = PosX + 375
    
                Contador = Contador - 1
                
            Loop
        
        Case 4
            '                                    **
            'Monta um conjunto de Blocos do tipo **, em "pctProx"
            
            'Cria uma variável que indicará o número de índice dos blocos
            Dim NumBlocoProx As Integer
            
            NumBlocoProx = 0
            
            'Define as posições iniciais dos dois primeiros Blocos
            PosX = 525
            PosY = 50
            
            ' Seleciona um número que se referenciará a um
            'estilo do Bloco (imagens armazenadas com valores
            'de "1" a "7" em "ImageListBlocos")
            EstiloBloco = Random(7)
            If EstiloBloco = 0 Then EstiloBloco = 1

            'Corrige os valores do contador
            Contador = 1

            'Executa o primeiro DO
            Do While Contador >= 0
           
                frmJogo.pctProxBloc(NumBlocoProx).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                frmJogo.pctProxBloc(NumBlocoProx).Top = PosY
                frmJogo.pctProxBloc(NumBlocoProx).Left = PosX
                frmJogo.pctProxBloc(NumBlocoProx).Visible = True
                
                PosX = PosX + 375
                
                NumBlocoProx = NumBlocoProx + 1
                Contador = Contador - 1
                
            Loop
            
            'Redefine as posições iniciais dos dois primeiros Blocos
            PosX = 525
            PosY = 425
            
            'Corrige os valores do contador
            Contador = 1
            
            'Executa o segundo DO
            Do While Contador >= 0
           
                frmJogo.pctProxBloc(NumBlocoProx).Picture = frmJogo.ImageListBlocos.ListImages(EstiloBloco).Picture
                frmJogo.pctProxBloc(NumBlocoProx).Top = PosY
                frmJogo.pctProxBloc(NumBlocoProx).Left = PosX
                frmJogo.pctProxBloc(NumBlocoProx).Visible = True
                
                PosX = PosX + 375
                
                NumBlocoProx = NumBlocoProx + 1
                Contador = Contador - 1
                
            Loop
    
    End Select
    
End Function

Function Random(Base_de_Calculo As Integer)

'=======================================================================
'  Esta Função, com base nos segundos do relógio interno do computador,
' seleciona uma das três possibilidades de cálculo de aleatoriedade e
' realiza a seleção de um número aleatório
'=======================================================================
    
    Dim vRnd As Integer
    
    Select Case (Base_de_Calculo * Second(Time))
    
        Case Is < 20
            vRnd = CInt((((((Second(Time) + Hour(Time) + Day(Date) + Month(Date) + Right(Second(Time) * Hour(Time), 1)) * 314158 + 1) Mod 10000) / 10000) * (Base_de_Calculo + 1)) Mod (Base_de_Calculo + 1))
             
        Case Is > 200
            vRnd = CInt((((((Second(Time) + Hour(Time) + Day(Date) + Month(Date) + Left(Month(Date) * Hour(Time) + Second(Time), 1)) * 314158 + 1) Mod 10000) / 10000) * (Base_de_Calculo + 1)) Mod (Base_de_Calculo + 1))
    
        Case Else
            vRnd = CInt((((((Second(Time) + Hour(Time) + Day(Date) + Month(Date)) * 314158 + 1) Mod 10000) / 10000) * (Base_de_Calculo + 1)) Mod (Base_de_Calculo + 1))
    
    End Select
    
    Random = vRnd

End Function

