<!--#include file="../../../Library/Common/micMainCon.asp" -->
<!--#include file="../../../Library/Common/funcoes.asp" -->
<!--#include file="websevice.asp" -->
<%
pedidoId = request.Cookies("Cielo")("processoId")

if pedidoId = "" then pedidoId = request.QueryString("pedidoId")

objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&request.Cookies("Cielo")("processoId")&"','IN Retorno / Sess�o Pedido "&pedidoId&"')")
objconn.execute("update emissaoProcesso set obs=obs+'|WS CIELO enviou retorno.' where id='"&pedidoId&"'")

set tidRS = objConn.execute("SELECT top 1 tid, pedidoId from dados_cielo WHERE pedidoId='"&pedidoId&"' order by id desc")

tid      = tidRS(0)
processo = tidRS(1)

	
sRequest = "<?xml version='1.0' encoding='UTF-8'?>" & _
"<requisicao-consulta id='1' versao='1.0.0' xmlns='http://ecommerce.cbmp.com.br'>" & _
"	<tid>"&tid&"</tid> " & _
"	<dados-ec> " & _
"		<numero>"&afiliacaoCielo&"</numero>" & _
"		<chave>"&chaveWS&"</chave>" & _
"	</dados-ec> " & _
"</requisicao-consulta> "

'response.write sRequest
'response.End()




		Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
		xmlhttp.SetTimeouts 30000, 30000, 70000, 70000
        xmlhttp.Open "POST",SoapServer,"false"
  		 xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
        xmlhttp.setRequestHeader "Content-Length", cStr(len(sRequest))

        xmlhttp.send "mensagem="&sRequest
        xml= xmlhttp.ResponseText 
		

	
    retorno_codigo_erro = pegaValorNode(xml,"erro//codigo")
    retorno_mensagem_erro = pegaValorNode(xml,"erro//mensagem")
	
if retorno_codigo_erro <> "" then
	response.write retorno_mensagem_erro
	  objconn.execute("update emissaoProcesso set obs=obs+'|WS CIELO enviou erro "&retorno_mensagem_erro&".' where id='"&request.Cookies("Cielo")("processoId")&"'")
objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&request.Cookies("Cielo")("processoId")&"','ERRO WS: "&retorno_mensagem_erro&"')")
objConn.EXECUTE("INSERT INTO dadosCielo_movimento (processoId, meio, acao, resposta) VALUES ('"&request.Cookies("Cielo")("processoId")&"','1','Retorno p�s processamento','"&retorno_mensagem_erro&"'")
	response.End()
end if
	
	objConn.EXECUTE("INSERT INTO dadosCielo_movimento (processoId, meio, acao, resposta) VALUES ('"&request.Cookies("Cielo")("processoId")&"','1','Retorno p�s processamento','Sem erro.')")

    valor = pegaValorNode(xml,"transacao//dados-pedido//numero")
    moeda = pegaValorNode(xml,"transacao//dados-pedido//valor")
    datahoraT = pegaValorNode(xml,"transacao//dados-pedido//data-hora")
    produto = pegaValorNode(xml,"transacao//forma-pagamento//produto")
    parcelas = pegaValorNode(xml,"transacao//forma-pagamento//pacelas")
    retorno_status = pegaValorNode(xml,"transacao//status")
	
	
	'autorizacao	Nó com dados da autorização caso tenha passado por essa etapa.
	autorizacaoCodigo = pegaValorNode(xml,"transacao//autorizacao/codigo")
	autorizacaoMsg = pegaValorNode(xml,"transacao//autorizacao/mensagem")
	autorizacaodatahoraT = pegaValorNode(xml,"transacao//autorizacao/data-hora")
	autorizacaoValor = pegaValorNode(xml,"transacao//autorizacao/valor")
	autorizacaoLr = pegaValorNode(xml,"transacao//autorizacao/lr")
	autorizacaoArp = pegaValorNode(xml,"transacao//autorizacao/arp")

	'captura	Nó com dados da captura caso tenha passado por essa etapa.
	capturaCodigo = pegaValorNode(xml,"transacao//captura/codigo")
	capturaMsg = pegaValorNode(xml,"transacao//captura/mensagem")
	capturadatahoraT = pegaValorNode(xml,"transacao//captura/data-hora")
	capturaValor = pegaValorNode(xml,"transacao//captura/valor")	
	
	
	if retorno_status = "9" then
    msgRetorno = pegaValorNode(xml,"transacao//cancelamento//mensagem")
	end if
	
	if retorno_status = "5" OR retorno_status = "4" OR retorno_status = "6" then
		msgRetorno = pegaValorNode(xml,"transacao//autorizacao//mensagem")
	end if
	
	if retorno_status = "4" OR retorno_status = "6" then
	autorizado = "S"
		response.write "Código Autorização: "& autorizacaoCodigo
		objConn.Execute("UPDATE emissaoProcesso set pgtoAprovado=2  WHERE id ="&pedidoId)
	else
		autorizado = "N"
		objConn.Execute("UPDATE emissaoProcesso set pgtoAprovado=1  WHERE id ="&pedidoId)
	end if


	SQL = "UPDATE dados_cielo SET" & _
		  " msgRetorno = '"&msgRetorno&"', " & _
		  " statusCielo = '"&retorno_status&"', " & _
		  " statusCIELOTXT = '"&statusCIELOTXT(retorno_status)&"', " & _
		  " autorizacaoCodigo = '"&autorizacaoCodigo&"', " & _
	      " autorizacaoMsg = '"&autorizacaoMsg&"', " & _
	      " autorizacaodatahora = '"&autorizacaodatahoraT&"', " & _
	      " autorizacaoValor = '"&autorizacaoValor&"', " & _
	      " autorizacaoLr = '"&autorizacaoLr&"', " & _
	      " autorizacaoArp = '"&autorizacaoArp&"', " & _
	      " capturaCodigo = '"&capturadatahoraT&"', " & _
	      " capturaMsg = '"&capturadatahoraT&"', " & _
	      " capturadatahora = '"&capturadatahoraT&"', " & _
	      " capturaValor = '"&capturaValor&"', " & _
		  " autorizado = '"&autorizado&"' " & _
		  "  WHERE pedidoId = '"&processo&"' and tid='"&tid&"'"
		  
	objConn.execute(SQL)
	


	SET processoRS = objConn.Execute("SELECT * from emissaoProcesso WHERE id="&pedidoId)
	produto = processoRS("produto")
	processoRS.close
	set processoRS = nothing
	


if autorizado = "S" then

		objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processo&"','Processo aprovado TID: "&tid&" COD: "&autorizacaoArp&" | redirecionado para finalizar')")
		objConn.close
		response.Redirect "autorizada.asp?processo="&pedidoId&"&pagamento=CC&ccMarca=VI&tid="&tid&"&autorizacao="&autorizacaoCodigo
else

	if retorno_status = "9" then
		objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processo&"','Processo nao finalizado | redirecionado para finalizar')")
		objConn.close
		response.Redirect "naofinalizada.asp?processoId="&pedidoId
	end if
	
	if retorno_status = "5" then
		objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processo&"','Processo nao autorizado TID: "&tid&" | redirecionado para finalizar')")
		objConn.close
		response.Redirect "naoautorizada.asp?processoId="&pedidoId
	end if

end if


%>

