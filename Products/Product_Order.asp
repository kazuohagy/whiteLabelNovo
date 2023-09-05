<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->

<%    
    modal = false
    address = "../../v2/Products/Plan_Comparative.asp?"

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
<style>
    .custom-card-height {
        height: 280px;
    }
     .vertical-center {
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        height: 100%;
    }
</style>
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
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center vertical-center"><h5>Assistência Médica e Hospitalar</h5>Os planos NEXT asseguram atendimento
           médico, hospitalar ou odontológico em caso de acidente ou enfermidade contraída durante a viagem, manifestados
            sob a forma de dor ou doença, incluindo opções que podem cobrir Covid-19.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center vertical-center"><h5>Telemedicina</h5>Opção Telemedicina para os viajantes que preferem uma consulta
           remota com um profissional médico para tratar de temas menores ou mesmo para conseguir uma intervenção primária veloz em quadros sensíveis e graves.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center vertical-center"><h5>Atraso e Extravio de Bagagem</h5>
            Mais do que ajudar na localização da bagagem extraviada, o seguro viagem NEXT cobre as suas despesas em relação e pertinência com o atraso ou extravio da bagagem.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center vertical-center"><h5>Cancelamento e Atraso de Voos</h5>Você receberá o reembolso de despesas necessárias com alimentação, locomoção e hotel nos casos
           documentados de atraso ou cancelamento de voo.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center vertical-center"><h5>Acidentes pessoais</h5>Você receberá a indenização prevista em sua apólice, caso sofra algum acidente durante a viagem, que seja passível da cobertura especificada no plano escolhido.</div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center vertical-center"><h5>Reembolso com Despesas Médicas</h5>Você receberá o reembolso das despesas realizadas com consultas médicas e medicamentos usados para
           o seu tratamento durante a viagem.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center  vertical-center"><h5>Atendimento 24h em Português</h5>
            Você tem a garantia de suporte e atendimento 24h, 7 dias por semana, no idioma que melhor convier.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center  vertical-center"><h5>Seguro viagem Europa</h5>
            Com os planos da NEXT, você assegura seu ingresso na Europa com planos completos e abrangentes que atendem a todas as exigências do Tratado de Schengen, além da garantia de suporte médico qualificado para dor ou doença em decorrência de enfermidade ou acidentes, até sua estabilização.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center  vertical-center"><h5>Seguro viagem EUA</h5>
            Contratando um plano NEXT para os EUA, você viaja com tranquilidade e proteção para todos os países da América do Norte, podendo, inclusive, cruzar fronteiras sem perder a cobertura. Você terá suporte médico de qualidade em qualquer situação de urgência e emergência, bastando acionar nossa central de emergência 24h pelo WhatsApp.

          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center  vertical-center"><h5>Seguro viagem Latam</h5>
            Com uma apólice da NEXT, você viaja por toda a América do Sul com a garantia de suporte e atendimento médico de qualidade nas situações de urgência e emergência médica que possam ocorrer durante a sua viagem internacional.
          </div>
        </div>
      </div>
      <div class="col-md-3 mx-2">
        <div class="card vantagem custom-card-height">
          <div class="card-body text-center  vertical-center"><h5>Seguro viagem Internacional</h5>
            Os planos NEXT atendem a todos os destinos mundiais, com raríssimas exceções. Todos os nossos planos oferecem apoio às situações de urgência e emergência médica, através de uma vasta rede credenciada em todo o mundo, com canais acessíveis, como o WhatsApp, sem burocracia e sem complicações.
          </div>
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