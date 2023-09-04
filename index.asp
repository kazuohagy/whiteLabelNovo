<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<!doctype html>
<html lang="pt-br">

<head>
  <title>Next Seguro Viagens</title>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" type="text/css" href="css/bootstrap.min.css">
  <link rel="stylesheet" type="text/css" href="css/estilos.css">
  <link rel="stylesheet" type="text/css" href="css/fontawesome/css/all.min.css">
  <link rel="stylesheet" type="text/css" href="css/slick-theme.css">
  <link rel="stylesheet" type="text/css" href="css/slick.css">
  <script src="js/bootstrap.bundle.min.js"></script>
  <script src="js/jquery-3.5.1.min.js"></script>
  <script src="js/jquery.mask.min.js"></script>
  <script src="js/slick.min.js"></script>
</head>

<body>
  <!--#include file ="./Components/Header.asp"-->
 
  <!-- Início banner -->
  <div class="imagem-capa" style="background-image: url('img/banner.png');"></div>
  <!-- Fim banner -->
  <div class="container">
    <!--#include file ="./Components/Price_Component.asp"-->   

    <div class="text-center mt-5">
      <h2>Vantagens do seguro viagem</h2>
    </div><br>

    <div class="row diff-cards slider mb-5">
      <div class="col-md-3 mx-2">
        <div class="card vantagem">
          <div class="card-body text-center">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus pharetra
            nisl eget tortor vestibulum semper. Maecenas dictum turpis purus, id finibus purus imperdiet sed. Morbi at
            augue vitae risus mattis lacinia.</div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem">
          <div class="card-body text-center">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus pharetra
            nisl eget tortor vestibulum semper. Maecenas dictum turpis purus, id finibus purus imperdiet sed. Morbi at
            augue vitae risus mattis lacinia.</div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem">
          <div class="card-body text-center">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus pharetra
            nisl eget tortor vestibulum semper. Maecenas dictum turpis purus, id finibus purus imperdiet sed. Morbi at
            augue vitae risus mattis lacinia.</div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem">
          <div class="card-body text-center">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus pharetra
            nisl eget tortor vestibulum semper. Maecenas dictum turpis purus, id finibus purus imperdiet sed. Morbi at
            augue vitae risus mattis lacinia.</div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem">
          <div class="card-body text-center">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus pharetra
            nisl eget tortor vestibulum semper. Maecenas dictum turpis purus, id finibus purus imperdiet sed. Morbi at
            augue vitae risus mattis lacinia.</div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem">
          <div class="card-body text-center">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus pharetra
            nisl eget tortor vestibulum semper. Maecenas dictum turpis purus, id finibus purus imperdiet sed. Morbi at
            augue vitae risus mattis lacinia.</div>
        </div>
      </div>
    </div>



    <script>
      $('.diff-cards').slick({
        infinite: true,
        arrows: false,
        dots: true,
        speed: 300,
        slidesToShow: 4,
        slidesToScroll: 4,
        autoplay: true,
        autoplaySpeed: 3000,
        responsive: [{
          breakpoint: 1024,
          settings: {
            slidesToShow: 3,
            slidesToScroll: 3,
            infinite: true,
            dots: true
          }
        },
        {
          breakpoint: 600,
          settings: {
            slidesToShow: 1,
            slidesToScroll: 1
          }
        },
        {
          breakpoint: 480,
          settings: {
            slidesToShow: 1,
            slidesToScroll: 1
          }
        }
          // You can unslick at a given breakpoint now by adding:
          // settings: "unslick"
          // instead of a settings object
        ]
      });
    </script>
  </div>

  <!--#include file ="./Components/Footer.asp"-->
  <script>
    $(document).ready(function () {
      $('.celular').mask('(00) 0000-00000');
      $('.telefone').mask('(00) 0000-00000');
      $('.data').mask('00/00/0000');
      $('.cep').mask('00.000-000');
      $('.cpf').mask('000.000.000-00');
    });
    function checkImageLink(imageUrl) {
        var img = new Image();

        img.onload = function() {
            document.getElementById("logo-marca").src = imageUrl;
        };

        img.onerror = function() {
            document.getElementById("logo-marca").src = "img/logo-parceiro.jpg";
        };

        img.src = imageUrl;
    }

    // Chame a função passando o link da imagem que você deseja verificar
    var imageUrl = "https://www.nextseguro.com.br/NextSeguroViagem/old/img/agtaut/<%=request.Cookies("wlabel")("revId")%>.jpg";
    checkImageLink(imageUrl);
    //teste
  </script>
</body>

</html>