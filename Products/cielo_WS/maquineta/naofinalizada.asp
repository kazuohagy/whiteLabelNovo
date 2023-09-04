<!--#include file="../../biblioteca/micMainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<%
processoId = request.QueryString("processoId")

'verLogado request.cookies("FCNET_MIC")("logado"),Request.ServerVariables("URL")
response.cookies("FCNET_MIC")("fluxoEmissao")=2
objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processoId&"','Exibido aviso de não finalizado.')")

objConn.execute("UPDATE emissaoPRocesso set pgtoAprovado='1' WHERE id="&processoId)


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<LINK href="../../css/main.css" rel=stylesheet type=text/css>
<title>Green Card Assist</title>
<meta http-equiv="Content-Type" content="text/html;charset=iso-8859-1" />
</head>

<body>
<table width="100" border="0" align="center" cellpadding="2" cellspacing="0" style="width:420px; height:336px;  border:solid 1px #000">
  <tr>
    <td width="27" rowspan="3" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="346" height="81" bgcolor="#FFFFFF"><img src="../../img/logoAffinity.png" border="0" width="409" height="159"/></td>
    <td width="35" rowspan="3" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
    <tr>
  	<td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td height="209" valign="top" bgcolor="#FFFFFF"><span class="alertaVermelho">Transa&ccedil;&atilde;o N&atilde;o Finalizada</span>.<br />
      <br />
      A transa&ccedil;&atilde;o n&atilde;o foi finalizada junto a operadora do cart&atilde;o.<br />
    <br /></td>
  </tr>
</table>
</body>
</html>
