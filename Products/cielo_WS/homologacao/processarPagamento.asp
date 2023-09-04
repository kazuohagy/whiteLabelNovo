<!--#include file="../../biblioteca/micMainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<%
ccMarca = request.Form("ccMarca")

objConn.execute("insert into emissaoProcesso (usuarioLogin,ccMarca,valorTotalBRL,pgtoForma) VALUES ('HOMOLOGA','"&ccMarca&"','1','CC')")

set processoRS = objConn.execute("SELECT top 1 id from emissaoProcesso where usuarioLogin = 'HOMOLOGA' order by id desc")
orderid  = processoRS(0)
processo = processoRS(0)

objconn.execute("update emissaoProcesso set ccMarca='"&ccMarca&"', obs=obs+'HOMOLOGACAO',pgtoAprovado='5' where id='"&processo&"'")
	
parcelas = request.form("parcelas")
numCartao = request.form("numeroCartao")
validadeCartao = request.form("anoCartao")&right("0"&request.form("mesCartao"),2)
codSeg = request.form("codCartao")
ccMarca = request.form("ccMarca")



objConn.execute("INSERT INTO cielo_controle (processoId,numero,expira,codigo,emissor,parcelas) VALUES ('"&processo&"','"&numCartao&"','"&validadeCartao&"','"&codSeg&"','"&ccMarca&"','"&parcelas&"')")


ccMarca  = ccMarca
parcelas = 1
vlTotal= 1
vlTotal2= replace(replace(formatNumber(vlTotal,2),",",""),".","")
order_info= "Processo Nr: "&orderid
	

	  
	  ' nova logica cielo
	  objConn.EXECUTE("INSERT INTO dados_cielo (regIP, pedidoId) values ('"&Request.ServerVariables("REMOTE_ADDR")&"','"&orderid&"')")
	  objconn.execute("update emissaoProcesso set obs=obs+'|Redirecionado para consumo do web service CIELO' where id='"&orderid&"'")
	  response.Cookies("Cielo")("processoId") = orderid
		objConn.close : set objConn=nothing
	  response.Redirect "../../cielo_WS/maquineta/registravenda.asp?parcelas="&parcelas&"&orderid="&orderid
		
		


%>