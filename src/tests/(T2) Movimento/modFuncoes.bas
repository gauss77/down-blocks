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

'On Error GoTo Erro

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
        Line Input #vLinhaTXT, linha
   
        ' Verifica se o texto da corrente linha do arquivo �
        'indica��o de nome de formul�rio
        If Left(linha, 1) = "*" Then
        
            Formulario = linha
            ' L� a pr�xima linha do arquivo texto (que indica
            'um controle do corrente formul�rio)
            Line Input #vLinhaTXT, linha
            
        End If
   
        'Verifica de qual formul�rio pertence o atual componente
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
                '   - vLabel: o r�tulo do controle
                '*Usa como base para separa��o da linha o sinal de
                'igual(=)
                vControle = Left(linha, InStr(1, linha, "=") - 1)
                vLabel = Right(linha, Len(linha) - InStr(1, linha, "="))
                
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
                vControle = Left(linha, InStr(1, linha, "=") - 1)
                vLabel = Right(linha, Len(linha) - InStr(1, linha, "="))

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
