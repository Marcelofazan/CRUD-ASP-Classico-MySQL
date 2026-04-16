<% 
    'On error resume next
    Dim conexao
    set conexao = server.createobject("adodb.connection")
    conexao.Open "DRIVER=MySQL ODBC 5.3 Unicode Driver;SERVER=127.0.0.1;PORT=3306;DATABASE=simplesweb;USER=root;PASSWORD=SUASENHA;OPTION=3;SSLMODE=disabled;"
        
    'if conexao.errors.count = 0 then 
    '    response.write "Ligação a Base com Sucesso"
    'else 
    '    response.write "Erro Problema"
    'end if
%>



