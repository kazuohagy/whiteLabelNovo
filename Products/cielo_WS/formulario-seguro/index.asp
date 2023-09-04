<!--#include file="../../../Library/Common/micMainCon.asp" -->
<!--#include file="../../../Library/Common/funcoes.asp" -->
<%
if request("orderK") <> "" then
	chave = request("orderK")
	set processoRS = objConn.execute("SELECT * FROM emissaoProcesso where chave='"&chave&"'")
else

	processoId = request("orderid")
	orderid    = request("orderid")
	
	
	set processoRS = objConn.execute("SELECT * FROM emissaoProcesso where id='"&processoId&"'")

end if
%>
<!DOCTYPE html>
<html lang="pt-br">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Formulario Seguro de Pagamento - VerifyBy CIELO</title>
  <meta name="viewport" content="initial-scale=1">
  <link rel="stylesheet" href="../../../CSS/card.css">
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.4.2/css/all.css">
  <script src="../../../JavaScript/jquery-3.5.1.min.js"></script>
  <script src="../../../CSS/bootstrap/js/bootstrap.min.js"></script>
  <script src="../../../CSS/credit-payment/card.js"></script>
  <script src="../../../CSS/credit-payment/jquery.card.js"></script>

  <script>

    $('form').card({
      // a selector or DOM element for the container
      // where you want the card to appear
      container: '.card-wrapper', // *required*

      formSelectors: {
        numberInput: 'input[name="number"]', // optional — default input[name="number"]
        expiryInput: 'input[name="expiry"]', // optional — default input[name="expiry"]
        cvcInput: 'input[name="cvc"]', // optional — default input[name="cvc"]
        nameInput: 'input[name="name"]' // optional - defaults input[name="name"]
      },

      width: 500, // optional — default 350px
      formatting: true, // optional - default true

      // Strings for translation - optional
      messages: {
        validDate: 'valid\ndate', // optional - default 'valid\nthru'
        monthYear: 'mm/yyyy', // optional - default 'month/year'
      },

      // Default placeholders for rendered fields - optional
      placeholders: {
        number: '•••• •••• •••• ••••',
        name: 'Full Name',
        expiry: '••/••',
        cvc: '•••'
      },


      masks: {
        cardNumber: '•' // optional - mask card number
      },

      // if true, will log helpful messages for setting up Card
      debug: false // optional - default false
    });
  </script>

  <script language="JavaScript">

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
      }

      if (document.getElementById('numeroCartao').value == '') {
        alert('\nInforme o numero do cartao');
        document.getElementById('numeroCartao').focus();
        document.getElementById('numeroCartao').style.backgroundColor = '#FF9999';
        return false;
      }

      if (document.getElementById('codCartao').value == '') {
        alert('\nInforme o codigo de seguranca');
        document.getElementById('codCartao').focus();
        document.getElementById('codCartao').style.backgroundColor = '#FF9999';
        return false;
      }

      if (document.getElementById('validade').value == '' && document.getElementById('validade').value == "undefined") {
        alert('\nInforme o ano do vencimento');
        document.getElementById('validade').focus();
        document.getElementById('validade').style.backgroundColor = '#FF9999';
        return false;
      }

      var dateCheck = document.getElementById('validade').value
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


  </script>


  <style>
    @import url('https://fonts.googleapis.com/css?family=Baloo+Bhaijaan|Ubuntu');

    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      font-family: 'Ubuntu', sans-serif;
    }

    body {
      background: #dadada;
      margin: 0 10px;
    }

    .payment {
      background: #f8f8f8; 
      max-width: 50vw;     
      margin: 80px auto;
      height: auto;
      padding: 35px;
      padding-top: 70px;
      border-radius: 5px;      
    }

    .payment h2 {
      text-align: center;
      letter-spacing: 2px;
      margin-bottom: 40px;
      color: #0d3c61;
    }

    .form .label {
      display: block;
      color: #555555;
      margin-bottom: 50px;
    }

    .input {
      padding: 13px 0px 13px 25px;      
      text-align: center;
      border: 2px solid #dddddd;
      border-radius: 5px;
      letter-spacing: 1px;
      word-spacing: 3px;
      outline: none;
      font-size: 16px;
      color: #555555;
    }

    .card-grp {
      display: flex;
      justify-content: space-between;
    }

    .card-item {
      width: 48%;
    }

    .space {
      margin-bottom: 20px;
    }

    .icon-relative {
      position: relative;
    }

    .icon-relative .fas,
    .icon-relative .far {
      position: absolute;
      bottom: 12px;
      left: 100px;
      font-size: 20px;
      color: #555555;
    }

    .btn {
      margin-top: 40px;
      background: #f78528;
      ;
      padding: 14px 40px;;
      text-align: center;
      color: #f8f8f8;
      border-radius: 5px;
      cursor: pointer;
      width: 100%;
    }

    .btnAdapt {
            font-family: 'geoMedium';
            border: 2px solid  #f78528;
            background-color: white;
            color: black;
            font-size: 16px;
            cursor: pointer;
        }

        .tev {
            font-family: 'geoMedium';
            border-color:  #f78528;
            color:  #f78528;
        }

        .tev:hover {
            font-family: 'geoMedium';
            background-color:   #f78528;
            color: white;
        }

    .centered {
      position: absolute;
      top: 30%;
      left: 50%;
      transform: translate(-50%, -50%);
      color: #fff;
      font-size: 52px;
      position: absolute;
      text-align: center;
    }

    @media screen and (max-width: 420px) {
      .card-grp {
        flex-direction: column;
      }

      .card-item {
        width: 100%;
        margin-bottom: 20px;
      }

      .btn {
        margin-top: 20px;
      }
    }

    
  </style>

</head>

<body>
  <header>
    <div class="content">
      <img class="header-img" src="../../../Images/planos/img-header.jpg" alt="" />
      <h1 class="centered">Fechamento de compra </h1>
    </div>
  </header>
  <div class="container">    
    <form action="../cielo_WS/maquineta/registraVenda.asp" method="post" id="form1" name="form1"
      onSubmit="document.getElementById('div_aguarde').style.display='aguardeeeeeeeeeeeeee';">      
      <div class="form-container active payment  ">                     
        <div class="card-wrapper"></div>
        <br>
        <label class="label">Nome impresso no Cartão:</label>  
        <div class="card space icon-relative">
          <input type="text" id="titular" class="input" placeholder="Nome no Cartão"
            onBlur="removeAspa(this)" name="name" required>
        </div>

        <label class="label">Numero do Cartão:</label>
        <div class="card space icon-relative">
          <input type="text" id="numeroCartao" class="input" placeholder="Número do Cartão"
            data-mask="0000 0000 0000 0000" onKeyPress="validate3(this.value)" name="number" required>
        </div>

        <label class="label">Data de Vencimento:</label>
        <div class="card space icon-relative">
          <input type="text" class="input" id="validade" placeholder="00 / 0000" name="expiry" required>
        </div>

        <label class="label">CVC:</label>
        <div class="card space icon-relative">
          <input type="text" id="codCartao" class="input" data-mask="000" placeholder="000" maxlength="3"
            onKeyPress="validate3(this.value)" name="cvc" required>
        </div>

        <label class="label">CCmarca:</label>
        <div class="card space icon-relative">
          <input type="text" id="ccMarca" class="input" value="" name="ccMarca" readonly>
        </div>

        <div class="card space icon-relative">
          <label class="compra">Valor da Compra:</label>
          <input type="text" class="input" value="R$ <%=formatNumber(processoRS("valorTotalBRL"),2)%>" readonly>
        </div>

        <div class="card space icon-relative">
          <div class="form-group">
            <label for="parcelas">Parcelamento:</label>
            <select class="form-control input" name="parcelas" id="parcelas" class="custom-select campoEmissao">
              <option value="1">Cartão à Vista</option>
              <% 
        
                If processoRS("valorTotalBRL") >=60  AND processoRS("valorTotalBRL") <=89.99 then
                  a = 2
                ElseIf processoRS("valorTotalBRL") >=90  AND processoRS("valorTotalBRL") <=119.99 then
                  a = 3
                ElseIf processoRS("valorTotalBRL") >=120  AND processoRS("valorTotalBRL") <=149.99 then
                  a = 4
                ElseIf processoRS("valorTotalBRL") >=150  AND processoRS("valorTotalBRL") <=999.99 then
                  a = 5
                ElseIf processoRS("valorTotalBRL") >=1000 then
                  a = 10
                else
                  a = 1
                end if
                
                FOR i=2 to a
        
              %>
              <option value="<%=i%>"><%=i%> parcelas</option>
              <%
                next
              %>
            </select>
          </div>
        </div>
        <input name="orderId" type="hidden" id="orderId" value="<%=processoRS("id")%>" />
        <input name="direto" type="hidden" id="direto" value="S" />      
        <button type="submit" class="btn i2Style btnAdapt tev" id="submit3" onclick="return confere_emissao();">teste</button>
      </div>
    </form>
  </div>

  <script>

    new Card({
      form: document.querySelector('form'),
      container: '.card-wrapper'
    });

  </script>

  <div id="div_aguarde"
    style="width:100%; height:600px; position:relative; top:0px; left:0px; text-align:center; font-size:24px; font-weight:700; color:#333; padding-top:200px; background:#CCC; position:absolute; display:none">
    Aguarde o processamento...
  </div>

</body>

</html>