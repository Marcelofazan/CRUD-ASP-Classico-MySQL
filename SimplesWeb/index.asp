<!DOCTYPE html>
<!--#include file="abreconexao.asp"-->
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
				            <h2 align="left">Simples Web</h2>
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