Attribute VB_Name = "FuncoesVener"
Option Explicit

'FuncoesVener
'Autor: Vener Fruet da Silveira
'Versão: 2023-05-16

Function importarArquivo(arquivo As String, delimitador As String, destino() As String) As Variant

'importarArquivo
'Função que lê o arquivo a ser importado

'Parâmetros:
'   arquivo = texto - caminho do arquivo a ser lido
'   delimitador = texto - delimitador de colunas
'   destino = matriz - destino dos dados lidos

'Autor: Vener Fruet da Silveira
'Versão: 2023-05-17

Dim matrixTabela() As String, matrixColunas() As String
Dim nrAquivo As Integer, txtLinha As String
Dim linhas As Integer, colunas As Integer, linha As Integer, coluna As Integer
    
    'Retorna a quantidade de linhas e colunas no arquivo
    linhas = contarLinhas(arquivo)
    colunas = contarColunas(arquivo, delimitador)
    
    'Define o tamanho da matriz tabela para a quatidade de linhas e a quantidade de colunas do arquivo
    ReDim destino(linhas, colunas)
    
    'Retorna nr de arquivo livre
    nrAquivo = FreeFile
    
    'Abre o arquivo
    Open arquivo For Input As nrAquivo
    
    'Percorre todo o arquivo até o final
    Do Until EOF(nrAquivo)
    
        'Lê uma linha por vez
        Line Input #nrAquivo, txtLinha
        
        'Divide a linha em colunas
        matrixColunas = explodeLinha(txtLinha, delimitador, colunas)
        
        'Adiciona dados das colunas a matrix tabela
        For coluna = 0 To colunas
            destino(linha, coluna) = matrixColunas(coluna)
        Next coluna
        
        linha = linha + 1
        
    Loop

    'Fecha o arquivo
    Close nrAquivo
    
    'Retorna matriz tabela
    importarArquivo = matrixTabela
    
End Function

Function explodeLinha(txtLinha As String, delimitador As String, colunas As Integer) As Variant
'explodeLinha
'Função que separa os dados de acordo com o delimitador

'Parâmetros:
'   txtLinha = texto - texto delimitado contendo os dados a serem separados
'   delimitador = texto - delimitador de colunas
'   colunas = número - inteiro de quantidade de colunas

'Autor: Vener Fruet da Silveira
'Versão: 2023-05-16

Dim matrixDados() As String, txtDado As String
Dim posInicio As Integer, posDelimitador As Integer, indice As Integer

    'Define a quantidade de colunas na matriz
    ReDim matrixDados(colunas)
    
    'Separa os dados em colunas na matriz
    Do
    
        'Retorna o primeiro delimitador
        posDelimitador = InStr(posDelimitador + 1, txtLinha, delimitador)
        
        If posDelimitador = 0 Then
            'Extrai a substring
            txtDado = Mid(txtLinha, posInicio + 1)
        Else
            'Extrai a substring
            txtDado = Mid(txtLinha, posInicio + 1, posDelimitador - posInicio - 1)
        End If
        
        'Popula a matriz
        matrixDados(indice) = txtDado
        
        'Define a posicição de inicio da próxima substring
        posInicio = posDelimitador
        
        'Define o próximo indice da matriz
        indice = indice + 1
    
    'Sai do loop se não existir mais delimitador
    Loop Until posDelimitador = 0
    
    'Retorna a matriz
    explodeLinha = matrixDados
    
End Function

Function contarColunas(arquivo As String, delimitador As String) As Integer

'contarColunas
'Função que retorna a quantidade de colunas baseado no delimitador

'Parâmetros:
'   arquivo = texto - caminho do arquivo a ser lido
'   delimitador = texto - delimitador de colunas

'Autor: Vener Fruet daSilveira
'Versão: 2023-05-17

Dim nrAquivo As Integer, txtLinha As String
Dim qtdeColunas As Integer, posDelimitador As Integer
    
    'Retorna nr de arquivo livre
    nrAquivo = FreeFile
    
    'Abre o arquivo
    Open arquivo For Input As nrAquivo
    
    'Lê a primeira para determinar a quantidade de colunas
    If Not EOF(nrAquivo) Then
        'Lê uma linha por vez
        Line Input #nrAquivo, txtLinha
    End If
    
    'Verifica a quantidade de delimitadores no texto
    Do
        posDelimitador = InStr(posDelimitador + 1, txtLinha, delimitador)
        qtdeColunas = qtdeColunas + 1
    Loop Until posDelimitador = 0
    
    'Fecha o arquivo
    Close nrAquivo
    
    'retorna a quantidade de colunas
    'deve-se subtrair 1 para garantir a quantidade correta de colunas
    contarColunas = qtdeColunas - 1
    
End Function

Function contarLinhas(arquivo) As Integer
'contarLinhas
'Função que retorna a quantidade de linhas no arquivo

'Parâmetros:
'   arquivo = texto - caminho do arquivo a ser lido

'Autor: Vener Fruet daSilveira
'Versão: 2023-05-17

Dim nrArquivo As Integer, linhas As Integer
Dim txt As String

    nrArquivo = FreeFile
    
    Open arquivo For Input As nrArquivo
    
    Do Until EOF(nrArquivo)
        Line Input #nrArquivo, txt
        linhas = linhas + 1
    Loop
    
    contarLinhas = linhas
    
    Close nrArquivo
    
End Function

Function exportarParaCSV(arquivo As String, listaDados As MSComctlLib.ListView, nomesColunas As Boolean) As Boolean
'Exporta dados para um arquivo delimitado por ponto e virgula (;)
'Parâmetros:
'   arquivo = texto - nome do arquivo a ser criado

'Autor: Vener Fruet da Silveira
'Versão: 2023-05-17

Dim cabecalho As ColumnHeaders
Dim itens As ListItems, item As ListItem
Dim txtValores As String
Dim linha As Integer, subItem As Integer
Dim nrArquivo As Integer, coluna As Integer

    'Define o objeto cabeçalho
    Set cabecalho = listaDados.ColumnHeaders
    
    'Define o objeto de itens da lista de dados
    Set itens = listaDados.ListItems
    
    'Retorna o arquivo livre
    nrArquivo = FreeFile
    
    'Inicia o tratamento de erro
    On Error GoTo trataErro
    
    'Abre o arquivo para inserção
    Open arquivo For Output As nrArquivo
    
    'Exportar cabeçalho?
    If nomesColunas Then
        For coluna = 1 To cabecalho.Count
            If coluna = 1 Then
                txtValores = cabecalho.item(coluna).Text
            Else
                txtValores = txtValores & "; " & cabecalho.item(coluna).Text
            End If
            
        Next coluna
        
        'Isere os dados no arquivo
        Print #nrArquivo, txtValores

    End If
    
    'Percorre a lista para exportar os dados
    For linha = 1 To itens.Count
    
        'Limpa a saida de dados
        txtValores = ""
        
        'Define o objeto item da lista de dados
        Set item = itens.item(linha)
        txtValores = item.Text
        
        'Percorre os subitens para completar os valores da tabelas
        For subItem = 1 To item.ListSubItems.Count
            txtValores = txtValores & "; " & item.ListSubItems.item(subItem).Text
        Next subItem
        
        'Isere os dados no arquivo
        Print #nrArquivo, txtValores
        
    Next linha
    
    'Mensagem de controle
    MsgBox "Exportação concluída", vbInformation, TITULO_PADRAO
    
    'Fecha o arquivo
    Close nrArquivo
    
    exportarParaCSV = True
    
    Exit Function
    
trataErro:
    'Caso exita um erro exibe a mensagem do erro
    MsgBox Err.Description, vbCritical, TITULO_PADRAO
    
    'Fecha o arquivo
    Close nrArquivo
    
    exportarParaCSV = False
    
End Function



