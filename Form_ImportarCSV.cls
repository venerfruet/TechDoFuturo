VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ImportarCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Formulário ImportarCSV
'Autor: vener Fruet daSilveira
'Versão: 2023-05-16

Dim matrixTabela() As String

Private Sub cmdAbrirCSV_Click()
Dim pathArquivo As String, testaArquivo As String
    
    'Solicita o caminho do arquivo a ser importado
    pathArquivo = InputBox("Informe o caminho do arquivo", TITULO_PADRAO, "C:\Users\VENER\Documents\TechDoFuturo1\COD_IBGE_SP.csv")
    If pathArquivo = "" Then Exit Sub 'Botão cancelar acionado
    
    'Testa se o arquivo existe
    testaArquivo = Dir(pathArquivo)
    If testaArquivo = "" Then
    
        'se arquivo inexistente
        MsgBox "Arquivo inexistente!", vbExclamation, TITULO_PADRAO
        Exit Sub
        
    Else
    
        'se arquivo existe
        importarArquivo pathArquivo, ";", matrixTabela()
        
        'Configura a lista de dados
        definirLista
        
        'Carrega lista
        carregarLista
        
        'Ativa o painel de opções
        Me.painelOpcoes.Enabled = True
        
    End If

End Sub

Private Sub definirLista()
Dim cabecalho As ColumnHeader
Dim coluna As Integer, colunas As Integer
Dim txtColuna As String

    'retorna a quantidade de colunas da matriz
    colunas = UBound(matrixTabela, 2)
    
    'Limpa o cabeçalho da lista
    Me.listaDados.ColumnHeaders.Clear
    
    'define o novo cabeçalho da lista
    For coluna = 0 To colunas
    
        'Retorna o objeto cabeçalho da lista
        Set cabecalho = Me.listaDados.ColumnHeaders.Add()
        
        'Verifica se deve definir a primeira linha como nome de colunas
        If Me.nomesColunas = True Then
            'Se sim a primeira é definida como nome de cabeçalho
            txtColuna = matrixTabela(0, coluna)
        Else
            'Se não o cabeçalho é definido como nomes de colunas sequenciais
            txtColuna = "Coluna " & coluna
        End If
        
        'define o texto e o tamanho das colunas do cabeçalho
        With cabecalho
            .Text = txtColuna
            .Width = Me.listaDados.Width / colunas
        End With
        
        'Esvazia o objeto cabeçalho
        Set cabecalho = Nothing
        
    Next coluna
    
End Sub

Private Sub carregarLista()
Dim item As ListItem
Dim linha As Integer, coluna As Integer
    
    'Esvazia a lista
    Me.listaDados.ListItems.Clear
    
    'Se primeira linha são os nomes das colunas então a lista começa no índice 1
    If Me.nomesColunas = True Then linha = 1
    
    'Percorre a matriz tabela
    For linha = linha To UBound(matrixTabela, 1)
        
        'Define um novo objeto de item da lista
        Set item = Me.listaDados.ListItems.Add
        'Define o texto d item, primeira coluna
        item.Text = matrixTabela(linha, 0)
        
        'insere as demais colunas no item da lista
        For coluna = 1 To UBound(matrixTabela, 2)
            item.ListSubItems.Add , , matrixTabela(linha, coluna)
        Next coluna
        
    Next linha
    
End Sub

Private Sub criarTabela_Click()
Dim dbs As Database
Dim nomeTabela As String, txtCampostabela As String, nomeCampo As String
Dim colunas As ColumnHeaders
Dim coluna As Integer

    'Solicita o nome para a nova tabela
    nomeTabela = InputBox("Digite o nome para a nova tabela", TITULO_PADRAO)
    'Caso não seja informado um nome termina o método
    If IsNull(nomeTabela) Or nomeTabela = "" Then Exit Sub
    
    'Define o objeto de banco de dados
    Set dbs = CurrentDb
    'define o objeto que contém os nomes das colunas
    Set colunas = Me.listaDados.ColumnHeaders
    
    'Percorre o objeto com o nomes das colunas para definir os nomes dos campos da nova tabela
    For coluna = 1 To colunas.Count
        
        'Retorna o nome para o campo
        nomeCampo = Trim(colunas.item(coluna))
        
        'Caso o nome do campo esteja vazio define com o nome campo em sequência
        IIf IsNull(nomeCampo) Or nomeCampo = "", "coluna" & coluna, nomeCampo
        
        'Caso a coluna seja maior que 1 concatena os textos separando por vigula
        If coluna = 1 Then
            txtCampostabela = "[" & nomeCampo & "] varchar(255)"
        Else
            txtCampostabela = txtCampostabela & ", [" & nomeCampo & "] varchar(255)"
        End If
        
    Next coluna
    
    'Inicia o tratamento de erro
    On Error GoTo trataErro
    
    'Executa o SQL DDL
    dbs.Execute "create table " & nomeTabela & "(" & txtCampostabela & ")"
    
    'Insere os dados na tabela
    inserirDados nomeTabela
    
    Set dbs = Nothing
    Exit Sub
    
trataErro:
    'Caso exita um erro exibe a mensagem do erro
    MsgBox Err.Description, vbCritical, TITULO_PADRAO
    
End Sub

Private Sub exportar_Click()
Dim pathArquivo As String, testaArquivo As String

    'Solicita o caminho do arquivo a ser importado
    pathArquivo = InputBox("Informe o caminho do arquivo", TITULO_PADRAO, "C:\Users\VENER\Documents\TechDoFuturo1\COD_IBGE_SP.csv")
    If pathArquivo = "" Then Exit Sub 'Botão cancelar acionado
    
    'Testa se o arquivo existe
    testaArquivo = Dir(pathArquivo)
    If testaArquivo <> "" Then
    
        'se arquivo inexistente
        MsgBox "Arquivo já existe!", vbExclamation, TITULO_PADRAO
        Exit Sub
        
    Else
    
        'Exporta os dados
        exportarParaCSV pathArquivo, Me.listaDados.Object, Me.nomesColunas
        
    End If
    
End Sub

Private Sub Form_Load()
    Me.painelOpcoes.Enabled = False
    Me.listaDados.ColumnHeaders.Clear
    Me.listaDados.ListItems.Clear
End Sub

Private Sub importarParaTabela_Click()
Dim nomeTabela As String

    'Abre a tela que lista os nomes das tabelas existentes no banco de dados
    DoCmd.OpenForm "ListarTabelas", acNormal, , , , acDialog
    
    On Error GoTo trataErro
    nomeTabela = Forms.item("ListarTabelas").listaTabelas
    
    inserirDados nomeTabela
    
    Exit Sub
    
trataErro:
    If Err.Number = 2450 Then
        Exit Sub
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub nomesColunas_AfterUpdate()
    'Evento disparado quando o campo que informa se
    'a primeira linha são nomes de colunas
    definirLista
    carregarLista
End Sub

Private Sub inserirDados(nomeTabela As String)
'Insere os dados na tabela
Dim dbs As Database
Dim itens As ListItems, item As ListItem
Dim txtValores As String
Dim linha As Integer, subItem As Integer, registros As Integer

    'Define o objeto de banco de dados
    Set dbs = CurrentDb
    
    'Define o objeto de itens da lista de dados
    Set itens = Me.listaDados.ListItems
    
    'Inicia o tratamento de erro
    On Error GoTo trataErro
    
    'Percorre a lista para inserir os dados na tabela
    For linha = 1 To itens.Count
    
        'Define o objeto item da lista de dados
        Set item = itens.item(linha)
        txtValores = """" & item.Text & """"
        
        'Percorre os subitens para completar os valores da tabelas
        For subItem = 1 To item.ListSubItems.Count
            txtValores = txtValores & ", """ & item.ListSubItems.item(subItem).Text & """"
        Next subItem
        
        'Executa o SQL DML
        dbs.Execute "insert into " & nomeTabela & " values(" & txtValores & ")"
        
        'Calcula registros adicionados
        registros = registros + dbs.RecordsAffected
        
    Next linha
    
    'Mensagem de controle
    MsgBox registros & " registros adicionados " & vbCrLf & _
                        " a tabela " & nomeTabela, vbInformation, TITULO_PADRAO
    
    'Exibe a tabela
    DoCmd.OpenTable nomeTabela, acViewNormal

    Set dbs = Nothing
    Exit Sub
    
trataErro:
    'Caso exita um erro exibe a mensagem do erro
    MsgBox Err.Description, vbCritical, TITULO_PADRAO
    
End Sub

