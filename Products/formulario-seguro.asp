<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<%
if request("orderK") <> "" then
chave = request("orderK")
set processoRS = objConn.execute("SELECT * FROM emissaoProcesso where chave='"&chave&"'")
else

processoId = request("orderid")
orderid    = request("orderid")


set processoRS = objConn.execute("SELECT * FROM emissaoProcesso where id='"&processoId&"'")

set emissaoTudoRS = objConn.execute("SELECT destino,planos.nome AS nomePlano, * FROM emissaoProcesso INNER JOIN planos on planoId = planos.id INNER JOIN viagem_destino ON viagem_destino.id = destino WHERE emissaoProcesso.id = '"&processoId&"'")
set cotacaoRS = objConn.execute("SELECT * FROM cotacao_reg WHERE id = '"&processoRS("cotacaoProcesso")&"'")

end if
nPax = processoRS("nPax")
valorTotalBRL = processoRS("valorTotalBRL")

dim mesInicio, mesFim, diaInicio, diaFim, anoInicio, anoFim
    'Recuperando as datas de inicio da viagem'
    diaInicio = LEFT(processoRS("dataInicio"),2)
    mesInicio = MID(processoRS("dataInicio"),4,2)
    anoInicio = RIGHT(processoRS("dataInicio"),4)

    'Recuperando as datas de fim da viagem'		
    diaFim = LEFT(processoRS("dataFim"),2)
    mesFim = MID(processoRS("dataFim"),4,2)
    anoFim = RIGHT(processoRS("dataFim"),4)

  'Converte mes em numero para nome do mes em ASP CLASSIC
  Function ConvertMes(mes)
      Select Case mes
          Case "01": ConvertMes = "Janeiro"
          Case "02": ConvertMes = "Fevereiro"
          Case "03": ConvertMes = "Março"
          Case "04": ConvertMes = "Abril"
          Case "05": ConvertMes = "Maio"
          Case "06": ConvertMes = "Junho"
          Case "07": ConvertMes = "Julho"
          Case "08": ConvertMes = "Agosto"
          Case "09": ConvertMes = "Setembro"
          Case "10": ConvertMes = "Outubro"
          Case "11": ConvertMes = "Novembro"
          Case "12": ConvertMes = "Dezembro"
      End Select
  End Function
  'Fim da função

%>
<!DOCTYPE html>
<html lang="pt-br">

<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Formulario Seguro de Pagamento - VerifyBy CIELO</title>
<meta name="viewport" content="initial-scale=1">
<style>
body {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
  margin: 0;
}
#footer {
  /* Estilos para o footer */
  margin-top: auto;
}
</style>
<!--#include file="../Components/HTML_Head2.asp" -->
</head>

<body id="affinity-page">
  <header>
    <!--#include file ="../Components/Header.asp"-->
  </header>
    <!-- Início banner -->
   <!--#include file ="../Components/Banner.asp"-->
    <!-- Fim banner -->
  <div class="container">
    <div class="text-center mt-5">
      <h2>Passo 3 - Realizar Pagamento</h2>
    </div>
    <ul class="progressbar d-flex justify-content-center mb-5">
      <li class="active">Seleção do Plano</li>
      <li class="active">Informar Dados</li>
      <li class="active">Realizar Pagamento</li>
    </ul>

    <div class="row">
      <div class="col-md-5">
        <div class="row">
          <div class="col-md-12 mb-3">
              <div class="card plano shadow">
                <div class="card-body py-4 px-4">
                  <h5 class="card-title"><i class="fas fa-list"></i> Resumo da Compra</h5>
                  <br>
                  <ul>
                    <li>Plano: <%= emissaoTudoRS("nomePlano")%></li>
                    <li>Origem: Brasil</li>
                    <li>Destino: <%= emissaoTudoRS("nome")%></li>
                    <li>Início da Vigência: <%=diaInicio%> de <%=convertMes(mesInicio)%> de <%=anoInicio%></li>
                    <li>Fim da Vigência: <%=diaFim%> de <%=convertMes(mesFim)%> de <%=anoFim%></li>
                    <li>Passageiros com até 64 anos: <%=cotacaoRS("nPax_Novo")%></li>
                    <li>Passageiros com mais de 65 anos: <%=cotacaoRS("nPax_Idoso")%></li>
                  </ul>
                </div>
              </div>
          </div>
          <div class="col-md-12 mb-3">
            <div class="card plano shadow">
              <div class="card-body py-4 px-4">
                <h5 class="card-title"><i class="fas fa-shopping-cart"></i> Valor Final</h5>
                <br>
                <h5 class="preco-plano">R$ <%=formatNumber(processoRS("valorTotalBRL"),2)%></h5>
                <span class="preco-por-pessoa">R$ <%=formatNumber(valorTotalBRL/nPax,2)%> por pessoa</span>
                <br>
                <b>ou em até 10x  sem juros de R$ <%=formatnumber(valorTotalBRL/10)%></b><br>
                <b><a id="parcelaSelec">10x</a> sem juros
                    de R$:  <a id="totalParcelado"><%=formatNumber(valorTotalBRL/10)%></a></b>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="col-md-7">
        <div class="card plano shadow">
          <div class="card-body py-4 px-4">
            <div class="card-wrapper"></div><br>
              <form action="cielo_API/registraVenda.asp" method="post" id="form1" name="form1" class="pagamento">
                <div class="row">
                  <div class="col-md-6 mb-3">
                    <input type="text" name="name" id="titular" class="form-control campo-customizado" onBlur="removeAspa(this)" placeholder="Nome impresso no cartão" required>
                  </div>
                  <div class="col-md-6 mb-3">
                    <input type="text" name="number" id="numeroCartao" class="form-control campo-customizado" placeholder="Número do cartão" onKeyPress="validate3(this.value)" required>
                  </div>
                  <div class="col-md-3 mb-3">
                    <input type="text" id="codCartao" onKeyPress="validate3(this.value)" name="cvc" class="form-control campo-customizado" placeholder="CVC" required>
                  </div>
                  <div class="col-md-3 mb-3">
                    <input maxlenght="7" type="text" id="validade" name="expiry" class="form-control campo-customizado" placeholder="MM/YYYY" required>
                  </div>
                  <div class="col-md-6 mb-3">
                    <select class="form-select campo-customizado" name="parcelas" id="parcelas" >
                      <option>Parcelas</option>
                        <% 
                            a = 10

                          FOR i=1 to a
                        %>
                        <option value="<%=i%>">
                            <%=i%> parcelas</option>
                        <%
                          next
                        %>
                    </select>
                  </div>
                </div>
                <input name="ccMarca" type="hidden"  id="ccMarca" class="input" value=""  readonly />
                <input name="orderId" type="hidden" id="orderId" value="<%=processoRS("id")%>" />
                <input name="direto" type="hidden" id="direto" value="S" />    
                <div class="text-end mb-4 mt-4"><button type="submit" id="submit3" onclick="return confere_emissao();" class="cta">EFETUAR PAGAMENTO</button></div>
              </form>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script src="../js/card.js"></script>
  <script>

  // Aqui vai comecar a mudanca dos valores das parcelas de acordo com o valor total
  function formatarReais(valor) {
            return new Intl.NumberFormat('br-BR').format(valor);
  }
  function parseCommaDecimalValue(value) {
    // Remove pontos de milhar e substitui a vírgula pelo ponto
    return parseFloat(value.replace(/\./g, "").replace(",", "."));
  }
  function changeParcela(parcela) {
    const parcelas = parcela == "Parcelas" ? 10 : parcela;

    var valorTotal = parseCommaDecimalValue("<%=processoRS("valorTotalBRL")%>");

    var totalParcelado = valorTotal / parcelas;
    document.getElementById("totalParcelado").innerHTML = formatarReais(totalParcelado.toFixed(2));
    document.getElementById("parcelaSelec").innerHTML = parcelas+"x";
  }
  document.getElementById("parcelas").addEventListener("change", function(){
    var selectedValue = this.value;
    changeParcela(selectedValue);
  });

  //fim da funcao de mudanca de parcelas
  new Card({
    form: document.querySelector('form'),
    container: '.card-wrapper'
  });

  
  $(document).ready(function () {
    $("#numeroCartao").change(verificaBandeira);
  });

  function verificaBandeira() {
    document.getElementById('ccMarca').value = Payment.fns.cardType($("#numeroCartao").val());
  }


  function validate3(objeto) {
    var keypress = event.keyCode;

    var sCaracteres = '0123456789';

    if (sCaracteres.indexOf(String.fromCharCode(keypress)) != -1) {
      event.returnValue = true;
    } else
      event.returnValue = false;
  }


  function removeAspa(campo) {
    campo.value = campo.value.replace("'", "").replace("'", "").replace("'", "").replace("'", "").replace("'", "").replace("'", "").replace("'", "")
  }


  function formatanumero(numero, decimais) {
    var num = parseFloat(numero);
    var result = num.toFixed(decimais);

    if (result == 'NaN') { result = '0.00' }

    return result;
  }

  function formatanumero2(numero, decimais) {
    var num = parseFloat(numero);
    var result = num.toFixed(decimais);

    if (result == 'NaN') { result = '0.00' }

    return result.replace(",", ".");
  }

  function confere_emissao() {

    if (document.getElementById('titular').value == '') {
      document.getElementById('titular').focus();
      document.getElementById('titular').style.backgroundColor = '#FF9999';
      alert('\nInforme o titular');
      return false;
    }else{
      document.getElementById('titular').style.backgroundColor = '#FFFFFF';
    }

    if (document.getElementById('ccMarca').value == '') {
      document.getElementById('numeroCartao').focus();
      document.getElementById('numeroCartao').style.backgroundColor = '#FF9999';
      alert('\nBandeira não reconhecida digite o número do cartão novamente');
      return false;
    }else{
      document.getElementById('numeroCartao').style.backgroundColor = '#FFFFFF';
    }

    if (document.getElementById('numeroCartao').value == '') {
      alert('\nInforme o numero do cartao');
      document.getElementById('numeroCartao').focus();
      document.getElementById('numeroCartao').style.backgroundColor = '#FF9999';
      return false;
    }else{
      document.getElementById('numeroCartao').style.backgroundColor = '#FFFFFF';
    }

    if (document.getElementById('codCartao').value == '') {
      alert('\nInforme o codigo de seguranca');
      document.getElementById('codCartao').focus();
      document.getElementById('codCartao').style.backgroundColor = '#FF9999';
      return false;
    }else{
      document.getElementById('codCartao').style.backgroundColor = '#FFFFFF';
    }

    if (document.getElementById('validade').value == '' && document.getElementById('validade').value == "undefined") {
      alert('\nInforme o ano do vencimento');
      document.getElementById('validade').focus();
      document.getElementById('validade').style.backgroundColor = '#FF9999';
      return false;
    }else{
      document.getElementById('validade').style.backgroundColor = '#FFFFFF';
    }
    //verifica se parcela esta preenchido
    if (document.getElementById('parcelas').value == 'Parcelas') {
      alert('\nInforme o numero de parcelas');
      document.getElementById('parcelas').focus();
  
      return false;
    }

    var dateCheck = document.getElementById('validade').value;
    var dateChekin = dateCheck.split('/');
    mm2 = <%=month(date) %>
    yy2 = <%=year(date) %>
    yy15 = <%=year(date) + 15 %>

    if (dateChekin[0] > 0 && dateChekin[0] <= 12 ) {
      if (dateChekin[1] > yy15 || dateChekin[1] < yy2) {
        alert('\nInforme um ano valido');
        document.getElementById('validade').focus();
        return false;
      }
    } else {
      alert('\nInforme um mês valido');
      document.getElementById('validade').focus();
      return false;
    }

  }
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
  <footer id="footer">
    <!--#include file ="../Components/Footer.asp"-->
  </footer>
</body>

</html>