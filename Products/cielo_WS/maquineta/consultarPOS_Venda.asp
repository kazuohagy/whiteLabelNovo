<!--#include file="../../biblioteca/MainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<!--#include file="websevice.asp" -->
<%
if request.QueryString("pedidoId") <> "" then
	pedidoId = request.QueryString("pedidoId")
	set tidRS = objConn.execute("SELECT top 1 tid, pedidoId from dados_cielo WHERE pedidoId='"&request.QueryString("pedidoId")&"' order by id desc")
end if

if request.QueryString("tid") <> "" then
	set tidRS = objConn.execute("SELECT top 1 tid, pedidoId from dados_cielo WHERE tid='"&request.QueryString("tid")&"' order by id desc")
end if

tid = tidRS(0)
processo = tidRS(1)
pedidoId = processo
	
	if 1 = 1 then
sRequest = "<?xml version='1.0' encoding='UTF-8'?>" & _
"<requisicao-consulta id='1' versao='1.0.0' xmlns='http://ecommerce.cbmp.com.br'>" & _
"	<tid>"&tid&"</tid> " & _
"	<dados-ec> " & _
"		<numero>"&afiliacaoCielo&"</numero>" & _
"		<chave>"&chaveWS&"</chave>" & _
"	</dados-ec> " & _
"</requisicao-consulta> "

response.write sRequest
response.End()


	'envio
   set objSrvHTTP = Server.CreateObject ("Msxml2.ServerXMLHTTP.4.0")
   objSrvHTTP.setTimeouts 5000, 60000, 10000, 10000 
   objSrvHTTP.setOption(2) = 13056
   objSrvHTTP.open "POST",SoapServer,false
   objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
   objSrvHTTP.send "mensagem="& sRequest
   
   'resposta
   xml = objSrvHTTP.responseXML.xml
   
	
'	Set objDynu = Server.Createobject("Dynu.HTTP") 
'	objDynu.SetURL SoapServer
'	objDynu.SetFormData "mensagem", sRequest
'	HTML = objDynu.PostURL()
'	Set objDynu = Nothing
'	xml = HTML

	
    retorno_codigo_erro = pegaValorNode(xml,"erro//codigo")
    retorno_mensagem_erro = pegaValorNode(xml,"erro//mensagem")
	
	if retorno_codigo_erro <> "" then
	'response.write retorno_mensagem_erro
	  objconn.execute("update emissaoProcesso set obs=obs+'|WS CIELO enviou erro "&retorno_mensagem_erro&".' where id='"&request.Cookies("micCielo")("processoId")&"'")
	end if

    valor = pegaValorNode(xml,"transacao//dados-pedido//numero")
    moeda = pegaValorNode(xml,"transacao//dados-pedido//valor")
    datahoraT = pegaValorNode(xml,"transacao//dados-pedido//data-hora")
    produto = pegaValorNode(xml,"transacao//forma-pagamento//produto")
    parcelas = pegaValorNode(xml,"transacao//forma-pagamento//pacelas")
    retorno_status = pegaValorNode(xml,"transacao//status")
	
	
'autorizacao	Nó com dados da autorização caso tenha passado por essa etapa.
	autorizacaoCodigo = pegaValorNode(xml,"transacao//autorizacao/codigo")
	autorizacaoMsg = pegaValorNode(xml,"transacao//autorizacao/mensagem")
	autorizacaoDataHora = pegaValorNode(xml,"transacao//autorizacao/data-hora")
	autorizacaoValor = pegaValorNode(xml,"transacao//autorizacao/valor")
	autorizacaoLr = pegaValorNode(xml,"transacao//autorizacao/lr")
	autorizacaoArp = pegaValorNode(xml,"transacao//autorizacao/arp")
	autorizacaoLr = pegaValorNode(xml,"transacao//autorizacao/lr")

'captura	Nó com dados da captura caso tenha passado por essa etapa.
	capturaCodigo = pegaValorNode(xml,"transacao//captura/codigo")
	capturaMsg = pegaValorNode(xml,"transacao//captura/mensagem")
	capturaDataHora = pegaValorNode(xml,"transacao//captura/data-hora")
	capturaValor = pegaValorNode(xml,"transacao//captura/valor")	
	
	
	if retorno_status = "9" then
    msgRetorno = pegaValorNode(xml,"transacao//cancelamento//mensagem")
	end if
	
	if retorno_status = "5" OR retorno_status = "4" OR retorno_status = "6" then
		msgRetorno = pegaValorNode(xml,"transacao//autorizacao//mensagem")
	end if
	
	if retorno_status = "4" OR retorno_status = "6" then
	autorizado = "S"
	'response.write "Código Autorização: "& autorizacaoCodigo
	objConn.Execute("UPDATE emissaoProcesso set pgtoAprovado=2 WHERE id ="&pedidoId)
	else
	autorizado = "N"
	objConn.Execute("UPDATE emissaoProcesso set pgtoAprovado=1 WHERE id ="&pedidoId)
	end if


	SQL = "UPDATE dados_cielo SET" & _
		  " msgRetorno = '"&msgRetorno&"', " & _
		  " statusCielo = '"&retorno_status&"', " & _
		  " statusCIELOTXT = '"&statusCIELOTXT(retorno_status)&"', " & _
		  " autorizacaoCodigo = '"&autorizacaoCodigo&"', " & _
	      " autorizacaoMsg = '"&autorizacaoMsg&"', " & _
	      " autorizacaoDataHora = '"&autorizacaoDataHora&"', " & _
	      " autorizacaoValor = '"&autorizacaoValor&"', " & _
	      " autorizacaoLr = '"&autorizacaoLr&"', " & _
	      " autorizacaoArp = '"&autorizacaoArp&"', " & _
	      " capturaCodigo = '"&capturaDataHora&"', " & _
	      " capturaMsg = '"&capturaDataHora&"', " & _
	      " capturaDataHora = '"&capturaDataHora&"', " & _
	      " capturaValor = '"&capturaValor&"', " & _
		  " autorizado = '"&autorizado&"', " & _
		  " lr = '"&autorizacaoLr&"' " & _
		  "  WHERE pedidoId = '"&processo&"'"
		  
	objConn.execute(SQL)
	end if
	
	  'objconn.execute("update emissaoProcesso set obs=obs+'|Direcionado para emissão de vouchers.' where id='"&pedidoId&"'")
	  
	SET processoRS = objConn.Execute("SELECT * from emissaoProcesso WHERE pgtoForma = 'CC' AND id="&pedidoId)

objConn.close

if request.QueryString("modelo") = "teste" then
	response.write xml
else
	response.Write("Dados atualizados com sucesso")
end if

%>
