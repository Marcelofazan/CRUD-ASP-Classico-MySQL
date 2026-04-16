<%@ Language="VBScript"%>
<!--#include file="abreconexao.asp"-->
<%
If Request.QueryString("acao") = "excluir" Then
    Dim idExcluir, sidExcluir
    idExcluir = Request.QueryString("id")
    sidExcluir = CStr(idExcluir)
   
    If sidExcluir <> "" Then
        Response.Write("<script>alert('O ID será excluído : " & sidExcluir & "');</script>")
        conexao.Execute("DELETE FROM pessoa WHERE Id = " & CInt(sidExcluir))
        Response.Redirect("pessoa_busca.asp")
    End If
End If
%>

<!DOCTYPE html>
<html lang="en-US">
   <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <title>SimplesWeb</title>
      <link rel="stylesheet" href="css/components.css">
      <link rel="stylesheet" href="css/icons.css">
      <link rel="stylesheet" href="css/responsee.css">
      <link rel="stylesheet" href="owl-carousel/owl.carousel.css">
      <link rel="stylesheet" href="owl-carousel/owl.theme.css"> 
      <link href='http://fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800&subset=latin,latin-ext' rel='stylesheet' type='text/css'>
      <script type="text/javascript" src="js/jquery-1.8.3.min.js"></script>
      <script type="text/javascript" src="js/jquery-ui.min.js"></script>    
   </head>
   <body class="size-960">
      <!-- HEADER -->
      <header>
         <div class="line">
            <div class="box">
               <div class="s-12 l-2">
                  <img class="s-5 l-12 center" src=""></div>
            </div>
         </div>
         <!-- TOP NAV -->  
         <div class="line">
            <nav class="margin-bottom">
               <p class="nav-text">Menu principal</p>
               <div class="top-nav s-12 l-10">
                  <ul>
                    <li><a href="index.asp">SimplesWeb</a></li>
                     <li>
                        <a>Cadastro de Pessoa</a>
                        <ul>
                           <li><a href="pessoa_formulario.asp">Formulário</a></li>
                           <li><a href="pessoa_busca.asp">Busca</a></li>
                        </ul>
                     </li>
                  </ul>
               </div>
               <div class="hide-s hide-m l-2">
                  <i class=""></i>
               </div>
            </nav>
         </div>
      </header>
      <section>
         <!-- CAROUSEL --> 
         <!-- HOME PAGE BLOCK -->      
         <!-- ASIDE NAV AND CONTENT -->
         <div class="line">
            <div class="box margin-bottom">
               <div class="margin">
               <!-- CONTENT -->
                  <article class="s-12 m-7 l-8">
                    <ul>
				         <h3 align="left">Lista de Pessoas</h3>
                     <%
                        'pega a pagina atual
                        pagina = request.ServerVariables("SCRIPT_NAME")
                        'numero de registros por pagina...
                        pageSize = 5

                        if(len(Request.QueryString("p")) = 0 )then
                           paginaAtual = 1
                        else
                           paginaAtual = CInt(Request.QueryString("p"))
                        end if

                        ' conta o numero de registros...
                        sql = "SELECT COUNT(*) AS total FROM pessoa"
                        set rs = conexao.execute(sql)

                        'total de registros
                        recordCount = Cint(rs("total"))

                        'calculamos o numero de paginas...

                        pageCount = Clng(recordCount / pageSize)

                        If pageCount < 1 then
                           pageCount = 1
                        end if

                        rs.Close()
                        Flag1 = INT(paginaAtual / pagesize)
                        PI = INT(Flag1 * pagesize)

                        IF PI = 0 THEN
                           PI = 1
                        END IF

                        PF = PI + pagesize - 1
                        vPesquisa = request.form("NomeFantasia")

                        ' selecionamos os registros...
                        sql = "SELECT * FROM pessoa WHERE NomeFantasia like '%" & vPesquisa & "%' AND RazaoSocial like '%" & vPesquisa & "%'" & " LIMIT " & (paginaAtual - 1) * pageSize & " , " & pageSize 

                        set rs = conexao.execute(sql)


                           Response.Write ("<table>")
                           Response.Write("<thead>")
                           Response.Write("<tr>")
                           Response.Write("<th>Editar</th>")
                           Response.Write("<th>Excluir</th>")
                           Response.Write("<th>Razao Social</th>")
                           Response.Write("<th>Nome Fantasia</th>")
                           Response.Write("<th>Cnpj</th>")
                           Response.Write("<th>Inscr. Est.</th>")
                           Response.Write("<th>Inscr. Mun.</th>")
                           Response.Write("<th>Data Cadastro</th>")
                           Response.Write("<th>Email</th>")
                           Response.Write("<th>Celular</th>")
                           Response.Write("<th>I.E. Dest.</th>")
                           Response.Write("</tr>")
                           Response.Write("</thead>")
                           
                        do while not rs.eof
                           ' aqui entra o q você quer exibir
                           Response.Write ("<tbody>")
                           Response.Write ("<table>")

                           Response.Write ("<tr>")
                           Response.Write("<td><a href='javascript:confirmarEditar(" & rs("id") & ")'>Editar</a></td>")
                           Response.Write("<td><a href='javascript:confirmarExclusao(" & rs("id") & ")'>Excluir</a></td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("RazaoSocial") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("NomeFantasia") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("Cnpj") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("InscrEstadual") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("InscrMunicipal") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("DataCadastro") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("Email") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("Celular") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("<td align='justify'>")
                           Response.Write rs("IeDestinatario") & "<br>"
                           Response.Write ("</td>")
                           Response.Write ("</tr>")
                              rs.MoveNext()
                           loop

                        Response.Write ("</tbody>")
                        Response.Write ("</table>")

                        Response.Write ("<B><strong> Página " & paginaAtual & " de " & pagecount & " </strong></B><br>")
                        ' cria os links de pagians...
                        IF CInt(paginaAtual) > 1 THEN
                        Response.Write "<a href='"&pagina&"?p=1'>Primeira</a> "
                        Else
                        Response.Write "<font color=""#ADADAD"">Primeira</font> "
                        END IF

                        if CInt(paginaAtual) > 1 then
                        Response.Write "<a href='"&pagina&"?p=" & paginaAtual - 1 &"'>Anterior</a> "
                        Else
                        Response.Write "<font color='#666666'>Anterior</font>  "
                        END IF

                        for i=1 to pageCount
                        Response.Write("<a href='"&pagina&"?p=" & i & "'>" & i & "</a> ")
                        next

                        IF (CInt(paginaAtual) < pagecount) THEN
                        IF CInt(PF) <> pagecount THEN
                        Response.Write "<a href='"&pagina&"?p=" & paginaAtual+1 & "'>Próxima</a> "
                        END IF
                        Else
                        Response.Write "<font color=""#ADADAD"">Próxima</font> "
                        END IF

                        IF (CInt(paginaAtual) <> pagecount) THEN
                        IF CInt(PF) <> pagecount THEN
                        Response.Write "<a href='"&pagina&"?p=" & pagecount & "'>Última</a> "
                        END IF
                        Else
                        Response.Write "<font color=""#ADADAD"">Última</font> "
                        END IF

                        rs.Close()
                     %>
		               <p align="justify"><a href="" class=""></a></p>
                     </ul>
                  </article>
                  <!-- ASIDE NAV -->
               </div>
            </div>
         </div>
      <!-- GALLERY CAROUSEL -->
      </section>
      <!-- FOOTER -->
      <footer class="line">
         <div class="box">
            <div class="s-12 l-6">
               <p>Exemplo ASP 2026.</p>
            </div>
            <div class="s-12 l-6">
               <a class="right" href="" title=""></a>
            </div>
         </div>
      </footer>                                                                   
      <script type="text/javascript" src="js/responsee.js"></script>               
      <script type="text/javascript" src="owl-carousel/owl.carousel.js"></script>
      
      <script type="text/javascript">

        function confirmarExclusao(id) {
            if (confirm('Tem certeza que deseja excluir a Pessoa ' + id + '?')) {
                window.location.href = 'pessoa_busca.asp?acao=excluir&id=' + id;
            }
        }

            function confirmarEditar(id) {
            if (confirm("Deseja realmente editar o registro " + id + "?")) {
                  window.location.href = 'pessoa_formulario.asp?acao=editar&id=' + id;
            }
         }

        jQuery(document).ready(function($) {
          var owl = $('#header-carousel');
          owl.owlCarousel({
            nav: true,
            dots: true,
            items: 1,
            loop: true,
            navText: ["&#xf007","&#xf006"],
            autoplay: true,
            autoplayTimeout: 3000
          });
          var owl = $('#gallery-carousel');
          owl.owlCarousel({
            nav: false,
            dots: true,
            loop: true,
            navText: ["&#xf007","&#xf006"],
            autoplay: true,
            autoplayTimeout: 3000,
            responsive: {
              0: {
                items: 1
              },
              769: {
                items: 2
              },
              960: {
                items: 4
              }
            }
          });
        })
      </script>
   </body>
</html>
<!--#include file="fechaconexao.asp"-->