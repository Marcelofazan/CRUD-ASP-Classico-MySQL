# CRUD-ASP-Classico-MySQL

Exemplo de criação CRUD em Asp Clássico

## Competências


- **HTML5** - Utilização de Template Responsivo [owlcarousel](https://owlcarousel2.github.io/OwlCarousel2/demos/demos.html)
- **MySQL** - Utilização do Driver Unicode 5.3 
- **ASP Clássico** - Utilização de Linguaguem VbScript 

## Execução da aplicação

Necessário instalar driver  [MySQL](mysql-connector-odbc-5.3.13-win32).
Necessário Habilitar Aplicativos de 32 Bits no servidor IIS do Windows. 
Para executar a aplicação é necessário rodar Script banco de dados. 

## String de conexão do banco

Modifique a string de conexão no arquivo **abreconexao.asp**, no trecho indicado:

```bash
...
    conexao.Open "DRIVER=MySQL ODBC 5.3 Unicode Driver;SERVER=127.0.0.1;PORT=3306;DATABASE=simplesweb;USER=root;PASSWORD=SuaSenha;OPTION=3;SSLMODE=disabled;"
...

```
O script para criação da tabela do exemplo encontra-se na pasta **Database**.

### Aqui está uma demonstração do Projeto

<img width="934" height="605" alt="ASP-Classico-CRUD-MySQL" src="https://github.com/user-attachments/assets/6869ba0c-13be-4604-9d4d-94c2070f0a7f" />


