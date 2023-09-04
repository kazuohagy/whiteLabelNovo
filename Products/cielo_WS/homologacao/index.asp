<!--#include file="../../biblioteca/micMainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<%

processoId = request.QueryString("processoId")

%>
<html>
<head>
<LINK href="../../css/main.css" rel=stylesheet type=text/css>
<title>Homologacao CIELO</title>
<meta http-equiv="Content-Type" content="text/html;charset=iso-8859-1" />

<script language="JavaScript">




function validate3(objeto) {
var keypress = event.keyCode; 
var campo = eval (objeto);

var sCaracteres = '0123456789';

if (sCaracteres.indexOf(String.fromCharCode(keypress))!=-1){
		event.returnValue = true;
}else
	event.returnValue = false;
}

 
function removeAspa(campo){
campo.value=campo.value.replace("'","").replace("'","").replace("'","").replace("'","").replace("'","").replace("'","").replace("'","")
}


function formatanumero(numero,decimais)
{ 
	var num = parseFloat(numero);
	var result = num.toFixed(decimais); 
				
	if(result=='NaN'){result='0.00'}
	
	return result; 
}

function formatanumero2(numero,decimais)
{ 
	var num = parseFloat(numero);
	var result = num.toFixed(decimais); 
				
	if(result=='NaN'){result='0.00'}
	
	return result.replace(",","."); 
}

function confere_emissao()
{ 
	
	if (document.getElementById('titular').value == ''){
		document.getElementById('titular').focus();
		document.getElementById('titular').style.backgroundColor = '#FF9999';
		alert('\nInforme o titular');
		return false;}
		
	if (document.getElementById('numeroCartao').value == ''){
		alert('\nInforme o numero do cartao');
		document.getElementById('numeroCartao').focus();
		document.getElementById('numeroCartao').style.backgroundColor = '#FF9999';
		return false;}

	if (document.getElementById('codCartao').value == ''){
		alert('\nInforme o codigo de seguranca');
		document.getElementById('codCartao').focus();
		document.getElementById('codCartao').style.backgroundColor = '#FF9999';
		return false;}

	if (document.getElementById('mesCartao').value == ''){
		alert('\nInforme o mes do vencimento');
		document.getElementById('mesCartao').focus();
		document.getElementById('mesCartao').style.backgroundColor = '#FF9999';
		return false;}

	if (document.getElementById('anoCartao').value == ''){
		alert('\nInforme o ano do vencimento');
		document.getElementById('anoCartao').focus();
		document.getElementById('anoCartao').style.backgroundColor = '#FF9999';
		return false;}

}



</script>

</head>

<body>
<form name="form1" method="post" action="processarPagamento.asp">
  <table width="800" border="0" align="center" cellpadding="2" cellspacing="0" style="width:800px; height:500px; background:url(../img/bkDialogo.jpg)">
    <tr>
      <td width="27" rowspan="3">&nbsp;</td>
      <td height="81" align="right">HOMOLOGA&Ccedil;&Acirc;O - COMPRA NAO VALIDA</td>
      <td width="35" rowspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top"><p><br>
        Compra de seguro viagem:<br>
          De: 01/12/2014 At&eacute; 02/12/2014<br>
          Vigencia: 1 dia<br>
          Valor
          : R$ 1,00</p>
        <table width="610" border="0" align="center" cellpadding="4" cellspacing="0" >
          <tr>
            <td style="padding:4px" width="33%" height="32" align="center"><span class="tdPreenche">
              <input name="ccMarca" type="radio" id="ccMarca3" value="VI" checked="checked" />
              </span><br />
              <img src="../img/visa.jpg" width="80" height="27" /></td>
            <td style="padding:4px" width="33%" align="center"><input type="radio" name="ccMarca" id="ccMarca4" value="MA" />
              <br />
              <img src="../img/master.png" width="60" height="38" /></td>
            <td style="padding:4px" width="33%" align="center"><input type="radio" name="ccMarca" id="ccMarca5" value="AM" />
              <br />
              <img src="../img/amex.jpg" width="50" height="50" /></td>
          </tr>
          <tr>
            <td style="padding:4px" height="32" colspan="3"><font class="tituloEmissao">Nome impresso no cart&atilde;o</font><br />
              <input name="titular" type="text" class="campoEmissao" id="titular" onBlur="maiuscula(this);removeAspa(this)" size="70" maxlength="255" style="padding:2px;text-transform:uppercase;"  /></td>
          </tr>
          <tr>
            <td style="padding:4px" height="32" colspan="3"><font class="tituloEmissao">N&uacute;mero do cart&atilde;o</FONT><br />
              <input name="numeroCartao" type="text" class="campoEmissao" id="numeroCartao" size="70" maxlength="255" style="padding:2px" onKeyPress="validate3(this.value)" /></td>
          </tr>
          <tr>
            <td style="padding:4px" height="32" valign="top"><font class="tituloEmissao">Validade do cart&atilde;o</font><br />
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><select name="mesCartao" class="campoEmissao" id="mesCartao" style="padding:2px">
                    <option value="">M&ecirc;s</option>
                    <%for i = 1 to 12%>
                    <option value="<%=i%>"><%=i%></option>
                    <%next%>
                  </select></td>
                  <td><select name="anoCartao" class="campoEmissao" id="anoCartao" style="padding:2px">
                    <option value="">Ano</option>
                    <%for i = year(date) to  year(date) + 7%>
                    <option value="<%=i%>"><%=i%></option>
                    <%next%>
                  </select></td>
                </tr>
              </table>
              <br /></td>
            <td colspan="2" valign="top" style="padding:4px"><font class="tituloEmissao">C&oacute;digo de seguran&ccedil;a</font><br />
              <input name="codCartao" type="text" class="campoEmissao" id="codCartao" size="10" maxlength="4" style="padding:2px" onKeyPress="validate3(this.value)"/></td>
          </tr>
          <tr>
            <td style="padding:4px" height="32"><font class="tituloEmissao">Valor da Compra:</font><br />
              R$ 1,00</td>
            <td height="32" colspan="2" style="padding:4px"><span class="tituloEmissao">Parcelamento</span><br />
              <select name="parcelas" id="parcelas" class="campoEmissao" style="padding:2px">
                <option value="1">Cart&atilde;o &agrave; Vista</option>
                <option value="2"> 2x </option>
              </select></td>
          </tr>
          <tr>
            <td style="padding:4px" height="32" align="center">&nbsp;</td>
            <td style="padding:4px" align="center"><input name="submit3" type="submit" class="i2Style" id="submit3" value="FINALIZAR" onClick="return confere_emissao();" />
              <input name="orderId" type="hidden" id="orderId" value="<%=request("orderId")%>" /></td>
            <td style="padding:4px" align="center">&nbsp;</td>
          </tr>
        </table>
</td>
    </tr>
   
  </table>
</form>

<div id="siteseal" style=" margin:0 auto; width:100%; text-align:center"><script type="text/javascript" src="https://seal.starfieldtech.com/getSeal?sealID=VaqFdMo1QQrvRH7TXttUBElr0So9cbq4oresTX6ZPmbO1FrlUbsq"></script></div>
</body>
</html>

