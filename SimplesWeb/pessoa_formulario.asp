	<%@ Language="VBScript"%>
   <!--#include file="abreconexao.asp"-->
   <%
      If Request.QueryString("acao") = "editar" Then
         Dim id, razaosocial, nomefantasia, cnpj, inscrestadual, inscrmunicipal, datacadastro, _
         email, celular, usuario, senha, iedestinatario

         idEditar = Request.QueryString("id")
         sidEditar = CStr(idEditar)
  
         If sidEditar <> "" Then
            set rs = conexao.Execute("SELECT * FROM pessoa WHERE Id = " & CInt(sidEditar))
            id = rs("Id")
            razaosocial = rs("RazaoSocial")
            nomefantasia = rs("NomeFantasia")
            cnpj = rs("Cnpj")
            inscrestadual = rs("InscrEstadual")
            inscrmunicipal = rs("InscrMunicipal")
            datacadastro = rs("DataCadastro")
            email = rs("Email")
            celular = rs("Celular")
            usuario = rs("Usuario")
            senha = rs("Senha")
            iedestinatario = rs("IeDestinatario")
            rs.Close()
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
						      <h3 align="left">Formulário de Pessoas</h3>
						      <br>
                        <form  action="pessoa_salvar.asp" method="GET" name="pessoa_formulario" id="pessoa_formulario" onSubmit="return validarCampos();"  class="customform">
                           <input type="hidden" name="txtid" id="txtid" value="<%=Request.QueryString("id")%>">
                           <div class="line">
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtRazaoSocial" id="txtRazaoSocial" placeholder="Digite a Razão Social" title="Digite a Razão Social" type="text"  maxlength="80" value="<%=razaosocial%>" />
                                 </div>
                              </div>
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtNomeFantasia" id="txtNomeFantasia" placeholder="Digite o Nome Fantasia" title="Digite o Nome Fantasia" type="text"  maxlength="80" value="<%=nomefantasia%>" />
                                 </div>
                              </div>
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtCnpj" id="txtCnpj"  placeholder="Digite o CNPJ" title="Digite o CNPJ" type="text"  maxlength="18" value="<%=cnpj%>" />
                                 </div>
                              </div>
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtInscrEstadual" id="txtInscrEstadual"  placeholder="Digite a Inscrição Estadual" title="Digite a Inscrição Estadual" type="text" maxlength="18" value="<%=inscrestadual%>" />
                                 </div>
                              </div>	
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtInscrMunicipal" id="txtInscrMunicipal"  placeholder="Digite a Inscrição Municipal" title="Digite a Inscrição Municipal" type="text" maxlength="18" value="<%=inscrmunicipal%>" />
                                 </div>
                              </div>	
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtDataCadastro" id="txtDataCadastro"  placeholder="Digite a Data Cadastro" title="Digite a Data Cadastro" type="text"  maxlength="18" value="<%=datacadastro%>" />
                                 </div>
                              </div>
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtEmail" id="txtEmail"  placeholder="Digite o Email" title="Digite o Email" type="text" maxlength="80" value="<%=email%>" />
                                 </div>
                              </div>	
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtCelular" id="txtCelular"  placeholder="Digite o Celular" title="Digite o Celular" type="text" maxlength="15" value="<%=celular%>" />
                                 </div>
                              </div>	
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtUsuario" id="txtUsuario"  placeholder="Digite o Usuario" title="Digite o Usuario" type="text"  maxlength="20" value="<%=usuario%>" />
                                 </div>
                              </div>
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtSenha" id="txtSenha"  placeholder="Digite a Senha" title="Digite a Senha" type="text" maxlength="12" value="<%=senha%>" />
                                 </div>
                              </div>	
                              <div class="margin">
                                 <div class="s-12 l-6">
                                    <input name="txtiedestinatario" id="txtiedestinatario"  placeholder="Digite o IE do Destinatário" title="Digite o IE do Destinatário" type="text" maxlength="15" value="<%=iedestinatario%>" />
                                 </div>
                              </div>	
                           </div>
                           <div class="margin">
                              <div class="s-12 l-2"><button type="submit">Gravar</button></div>
                              <div class="s-12 l-2"><button type="button" onclick="window.location.href='pessoa_busca.asp';">Cancelar</button></div>
                           </div>
                        </form>
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