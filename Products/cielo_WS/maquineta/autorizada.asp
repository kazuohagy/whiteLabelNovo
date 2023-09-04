<!--#include file="../../../Library/Common/micMainCon.asp" -->
<!--#include file="../../../Library/Common/funcoes.asp" -->
<%
processoId = request.QueryString("processo")

objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processoId&"','Autorizada: Redirecionando para recibo.')")
objConn.execute("UPDATE emissaoProcesso set pgtoAprovado='2' WHERE id="&processoId)

set processoRS = objConn.execute("SELECT * FROM emissaoPRocesso WHERE id="&processoId)
'set cieloRS = objConn.execute("SELECT * FROM dados_cielo where pedidoId='"&processoRS("id")&"'")

if UCASE(processoRS("usuarioLogin")) <> "HOMOLOGA" then

	SELECT CASE  processoRS("B2W")
	
		CASE 0: response.Redirect "../../emitir.asp?AcaoSubmit=0&pg=CC&processo="&processoId
		CASE 1: response.Redirect "../../emitir.asp?AcaoSubmit=0&pg=CC&processo="&processoId
		CASE 2: response.Redirect "http://www.goaffinity.com.br/emissao/emitir.asp?processo="&processoId & "&chave=" &processoRS("chave")
		CASE ELSE: response.Redirect "../../emitir.asp?AcaoSubmit=0&pg=CC&processo="&processoId
		
	END SELECT

'if processoRS("B2W") <> "2" then

'	response.Redirect "../../emissaoCliente/emitir.asp?AcaoSubmit=0&pg=CC&processo="&processoId

'else

'	response.Redirect "http://www.goaffinity.com.br/emissao/emitir.asp?processo="&processoId & "&chave=" &processoRS("chave")

'end if

else



%>
<html>
<head>      
<title>Autorizada</title>
<meta http-equiv="Content-Type" content="text/html;charset=iso-8859-1" />
</head>

<body>
<table width="800" border="0" align="center" cellpadding="2" cellspacing="0" style="width:800px; height:500px; background:url(../img/bkDialogo.jpg)">
  <tr>
    <td width="27" rowspan="3">&nbsp;</td>
    <td width="346" height="81">&nbsp;</td>
    <td width="35" rowspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="209" valign="top"><p><span class="alertaVermelho">Transaçãoo  Autorizada</span>.<br />
        <br>
      Voce receberá seu voucher no seu e-mail indicado no cadastro.</p>
      <p>Pedido: <%=request("processo")%><br>
        TID: <%=request("tid")%><br>
        Código de Autorização: <%=request("autorizacao")%></p>
<p>&nbsp;</p>
<a href="https://www.affinityseguro.com.br/Home/Index.asp" class="button">Retorne para página inicial</a>
</td>
  </tr>
</table>
</body>
</html>
<%end if%>