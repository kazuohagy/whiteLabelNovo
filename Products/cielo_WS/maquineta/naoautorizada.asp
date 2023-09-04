<!--#include file="../../../Library/Common/micMainCon.asp" -->
<!--#include file="../../../Library/Common/funcoes.asp" -->
<%
processoId = request.QueryString("processoId")

'verLogado request.cookies("FCNET_MIC")("logado"),Request.ServerVariables("URL")
response.cookies("FCNET_MIC")("fluxoEmissao")=2
objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processoId&"','Exibido aviso de n�o autorizada')")

objConn.execute("UPDATE emissaoPRocesso set pgtoAprovado='1' WHERE id="&processoId)

set processoRS = objConn.execute("SELECT * FROM emissaoPRocesso WHERE id="&processoId)


set cieloRS = objConn.execute("SELECT * FROM dados_cielo where pedidoId='"&processoRS("id")&"'")



%>
<html>
<head>
<LINK href="../../css/main.css" rel=stylesheet type=text/css>
<title>Affinity Assistencia</title>
<meta http-equiv="Content-Type" content="text/html;charset=iso-8859-1" />
<style>
td {
  padding:4px;
  font-family:Arial, Helvetica, sans-serif;
  font-weight:normal;
  font-size:12px;
}

a{color: #093;}
a:hover {color: #666;}

#geral {
position:absolute;
top:50%;
left:50%;
width:500px;
height:400px;
margin-left:-250px;
margin-top:-300px;
text-align:center;
}

.alertaVermelho {
	color:#900;
	font-family:Arial, Helvetica, sans-serif;
	font-size:16px;
	font-weight:bold;
}

</style>

</head>

<body>
<div id="geral" >
<table width="100" border="0" align="center" cellpadding="2" cellspacing="0" style="width:500; height:336px;  border:solid 1px #000">
  <tr>
    <td width="27" rowspan="3" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="346" height="81" align="center" bgcolor="#FFFFFF"><img src="https://www.affinityseguro.com.br/imgs/logo_topo.png" border="0" width="280" height="116"/></td>
    <td width="35" rowspan="3" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
    <tr>
  	<td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td height="209" valign="top" bgcolor="#FFFFFF">
    <p><span class="alertaVermelho">(<%=cieloRS("autorizacaoLr")%>) Transa&ccedil;&atilde;o n&atilde;o Autorizada.</span><br />
      <br />
      A transa&ccedil;&atilde;o n&atilde;o foi autorizada pela operadora do cart&atilde;o.<br />
      <br />
Pedido: <%=processoRS("id")%><br>
TID: <%=cieloRS("tid")%><br>

<%if not isnull(cieloRS("autorizacaoLr")) then%>

Status da transacao: 

<%
	set lrRS = objConn.execute("SELECT * from cielo_lr where lrCodigo='"&cieloRS("autorizacaoLr")&"'")
	if not lrRS.eof then response.write lrRS("descricao")
	set lrRS = nothing
end if
%>

<a href="https://www.affinityseguro.com.br/Home/Index.asp" class="button">Retorne para página inicial</a>

</p>
<%if 1 <> 1 then%>
    <p>&nbsp;</p>
    Op&ccedil;&otilde;es disponiveis:
    <table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="8%">&nbsp;</td>
        <td width="92%"><a href="../../emissao/pagamentoCC.asp?processo=<%=processoRS("id")%>&parcelas=1&orderid=<%=processoRS("id")%>&idAge=<%=processoRS("clienteId")%>&planoId=<%=processoRS("planoId")%>">Informar outro cart&atilde;o de cr&eacute;dito</a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><a href="../../emissao/emissao_entra.asp?carregarReserva=1&processo=<%=processoRS("id")%>">Recarregar processo e alterar forma de pagamento</a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><a href="../../emitidos/indexReserva.asp">Ir para processos reservados</a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><a href="../../emissao/viagem_entra.asp?n=1">Iniciar outra emiss&atilde;o</a></td>
      </tr>
    </table>
    <p>&nbsp;</p>


    </td>
  </tr>
 
</table>
<%end if%>
</div>
</body>
</html>
