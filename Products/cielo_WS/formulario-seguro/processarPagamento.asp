<!--#include file="../../biblioteca/micMainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<%
ccMarca = request.Form("ccMarca")

orderid  = request.Form("orderid")
processo = orderid

objconn.execute("update emissaoProcesso set ccMarca='"&ccMarca&"', obs=obs+'FORMULARIO CC',pgtoAprovado='5' where id='"&processo&"'")
	
parcelas = request.form("parcelas")
numCartao = request.form("numeroCartao")
validadeCartao = request.form("anoCartao")&right("0"&request.form("mesCartao"),2)
codSeg = request.form("codCartao")
ccMarca = request.form("ccMarca")
titular = request.form("titular")


objConn.execute("INSERT INTO cielo_controle (processoId,numero,expira,codigo,emissor,parcelas,titular) VALUES ('"&processo&"','"&numCartao&"','"&validadeCartao&"','"&codSeg&"','"&ccMarca&"','"&parcelas&"','"&titular&"')")


ccMarca  = ccMarca
parcelas = parcelas
vlTotal2= replace(replace(formatNumber(vlTotal,2),",",""),".","")
vlTotal= vlTotal2
order_info= "Processo Nr: "&orderid
	

	  
	  ' nova logica cielo
	  objConn.EXECUTE("INSERT INTO dados_cielo (regIP, pedidoId) values ('"&Request.ServerVariables("REMOTE_ADDR")&"','"&orderid&"')")
	  objconn.execute("update emissaoProcesso set obs=obs+'|Redirecionado para consumo do web service CIELO' where id='"&orderid&"'")
	  response.Cookies("Cielo")("processoId") = orderid
		objConn.close : set objConn=nothing
	  response.Redirect "../../cielo_WS/maquineta/registravenda.asp?parcelas="&parcelas&"&orderid="&orderid
		
		


%>