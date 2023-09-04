
<%
set tempRS = objConn.execute("SELECT * from goAffinity_afiliado where clienteId=" & request.Cookies("wlabel")("revId"))

if tempRS("bannerNome") <> "" then
  banner = "https://seguroviagemnext.com.br/Products/uploads/"&tempRS("bannerNome")&".jpg"
else
  banner = "https://seguroviagemnext.com.br/img/banner.png"
end if


%>
<div class="imagem-capa" id="imagemCapa" style="background-image: url('<%=banner%>'); display: flex; align-items: center; justify-content: center; flex-direction: column; text-align: center;">
  <h2 style="color: white;"><%=Request.Cookies("wlabel")("fantasia")%></h2>
  <h4 style="color: white;"><%=Request.Cookies("wlabel")("emailAtendimento")%></h4>
</div>
<script>
    function checkImageLink(logoUrl) {
      //alterar logo e banner caso não exista
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