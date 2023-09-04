<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->

<%    
    modal = false
    address = "../../Products/Plan_Comparative.asp?"

    siteNome = request.QueryString("siteNome")
	
    if siteNome <> "" then    
        set afiliadoRS   = objConn.execute("select clienteContato.id AS contatoId,goAffinity_afiliado. * from goAffinity_afiliado INNER JOIN clienteContato ON goAffinity_afiliado.xmlLogin = clienteContato.login where siteNome = '"&siteNome&"'")        
    
        if not afiliadoRS.eof then
            set clienteRS    = objConn.execute("select * from cadCliente where id = '"&afiliadoRS("clienteID")&"'")

            if clienteRS("ativo") <> "1" then
                response.Redirect "http://"& COMPANY_URL
            end if
        
            response.Cookies("wlabel")("revId")		    	= clienteRS("id")
            response.Cookies("wlabel")("razao")  			= clienteRS("razao")
            response.Cookies("wlabel")("fantasia")  		= clienteRS("fantasia")
            response.Cookies("wlabel")("CNPJ")   			= clienteRS("cnpj")
            response.Cookies("wlabel")("fone")				= afiliadoRS("foneAtendimento")
            response.Cookies("wlabel")("email")				= afiliadoRS("emailAtendimento")
            response.Cookies("wlabel")("xml_Id")			= afiliadoRS("contatoId")
            response.Cookies("wlabel")("xml_A")				= afiliadoRS("xmlLogin")
            response.Cookies("wlabel")("xml_B")				= afiliadoRS("xmlSenha")
            response.Cookies("wlabel")("siteNome")			= afiliadoRS("siteNome")
            response.Cookies("wlabel")("emailAtendimento")	= afiliadoRS("emailAtendimento")
            'response.Cookies("wlabel")("estilo")			= afiliadoRS("css")                
        else        
            response.Redirect "http://"& COMPANY_URL
        end if    
            
        set afiliadoRS = nothing
        set clienteRS = nothing
    else        
        response.Redirect "http://"& COMPANY_URL  
    end if    
%>
<!doctype html>
<html lang="pt-br">

<head>
  <title><%=COMPANY_NAME%></title>
<!--#include file="../Components/HTML_Head.asp" -->
</head>

<body>
  <!--#include file ="../Components/Header.asp"-->
  
  <!-- Início banner -->
    <!--#include file ="../Components/Banner.asp"-->
  <!-- Fim banner -->
  <div class="container">
    <!--#include file ="../Components/Price_Component.asp"-->   

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
  <footer id="footer">
    <!--#include file ="../Components/Footer.asp"-->
  </footer>
  <script>
    $(document).ready(function () {
      $('.celular').mask('(00) 0000-00000');
      $('.telefone').mask('(00) 0000-00000');
      $('.data').mask('00/00/0000');
      $('.cep').mask('00.000-000');
      $('.cpf').mask('000.000.000-00');
    });
    function checkImageLink(logoUrl) {
      //alterar logo  caso não exista
      var img = new Image();
      img.onload = function() {
        document.getElementById("logo-marca").src = logoUrl;
        
      };

      img.onerror = function() {
        document.getElementById("logo-marca").src = "./img/logo-parceiro.jpg";
      };

      img.src = logoUrl;
    }

    // Chame a função passando o link da imagem que você deseja verificar
    document.addEventListener("DOMContentLoaded", function() {
      var logoUrl = "https://www.nextseguro.com.br/NextSeguroViagem/old/img/agtaut/<%=request.Cookies("wlabel")("revId")%>.jpg";
     
      checkImageLink(logoUrl);
    });

   
  </script>
</body>

</html>