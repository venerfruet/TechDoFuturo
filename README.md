# Tech do Futuro 1º edição
## Repositório para entrega do desafio de código do Projeto Tech do Futuro Paschoalotto

### Do programa CSVImport:
Foi desenvolvido para mostrar minhas habilidades no desenvolvimento em espiral e exemplificar os métodos para manipulação de arquivos de texto, não foi usada nenhuma API de importação do arquivo CSV para o banco de dados.  

Todos os métodos contidos neste programa foram desenvolvidos por mim.  

Prezei em demonstrar minhas habilidades no desenvolvimento de códigos utilizando conceitos como abstração de código, utilização de métodos nativos do VB/VBA para a leitura e gravação de arquivos de textos, manipulação de matrizes, tratamento de erros, utilização de objetos visuais para as telas e instrções SQL para DDL e DML na manipulação do banco de dados.  

Espero que agrade.

---
### Se não possuir o Access 32 bits.  
Faça o download do runtime do **MsAccess de 32 bits** no site da Microsoft.  
https://www.microsoft.com/pt-br/download/details.aspx?id=10910  

Ao abrir o programa com o runtime irá aparecer a tela abaixo.  
![image](https://github.com/venerfruet/TechDoFuturo/assets/105865020/c24746db-f94e-4e19-88dd-634437f962a7)
 #### CLIQUE EM ABRIR  
 
Caso você queira criar um local de confiabilidade definitivo será necessário alterar o registro do Windows, como mostrado abaixo:  

[HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations\Location(n)]  

![image](https://github.com/venerfruet/TechDoFuturo/assets/105865020/df562ddc-cc23-4b08-9926-aef040434718)
Onde cada chave Location é um local de confiabilidade.  

### Arquivos do repositório:

#### CSVImport.mdb  
  Para importar qualquer arquivo delimitado por ponto e vírgula (;).  
  Programa de código aberto que roda no Ms Access.  
  Para rodar o programa apartir do Ms Office 2010 este arquivo deverá ser definido como confiável.  
  Para acessar a estrutura basta abrir o programa com a tecla "SHIFT" pressionada.

#### COD_IBGE_SP.csv  
  Arquivos de dados de amostra.

#### Form_ImportarCSV.cls  
  Código da tela principal do programa

#### Form_ListarTabelas.cls  
  Código da tela auxiliar para exibir as tabelas existentes no banco de dados.

#### Ambiente.bas  
  Código do módulo de definição de variáveis de ambiente.

#### FuncoesVener.bas  
  Código do módulo das funções de importação e exportação de arquivos de texto.

---

#### Tela principal:


![image](https://github.com/venerfruet/TechDoFuturo/assets/105865020/e95181a1-5ce3-4b59-9085-323332f73b50)
![image](https://github.com/venerfruet/TechDoFuturo/assets/105865020/944f31ec-2db9-4da5-bc30-cf49f958829d)
![image](https://github.com/venerfruet/TechDoFuturo/assets/105865020/570671ab-080e-4b78-921a-57acf0e9ec0e)
