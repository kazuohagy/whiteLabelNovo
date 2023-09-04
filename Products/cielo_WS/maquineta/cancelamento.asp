<!--#include file="../../biblioteca/micMainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<!--#include file="websevice.asp" -->
<%
processo = request.QueryString("pedidoId")

	sql= "select * from dados_cielo where pedidoId="&processo
	Set rs= objConn.Execute(sql)
	if rs.eof then response.write "Erro carregando dados do database processo total." : response.end

	
	
	numeroPedido = processo
	valor = vlTotal
	'datahoraT = YEAR(now) & "-" & RIGHT("0"&MONTH(now),2) & "-"  & RIGHT("0"&DAY(now),2) & "T" & TIME
	

sRequest = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf & _
"<requisicao-cancelamento id='1' versao='1.0.0' xmlns='http://ecommerce.cbmp.com.br'>" & vbcrlf & _
"	<tid>"&rs("tid")&"</tid>" & vbcrlf & _
"	<dados-ec>" & vbcrlf & _
"		<numero>"&afiliacaoCielo&"</numero>" & vbcrlf & _
"		<chave>"&chaveWS&"</chave>" & vbcrlf & _
"	</dados-ec>" & vbcrlf & _
"</requisicao-cancelamento>" 

   

if 1 <> 1 then ' script parou de funcionar por porblemas no uol
	'envio
   set objSrvHTTP = Server.CreateObject ("Msxml2.ServerXMLHTTP.6.0")
	'Tipos resolveTimeout,conectaTimeout,sendTimeout,receiveTimeout
	objSrvHTTP.setTimeouts 300000, 300000, 300000, 300000
   objSrvHTTP.setOption(2) = 13056
   objSrvHTTP.open "POST",SoapServer,false
   objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
   objSrvHTTP.send "mensagem="&sRequest
   
If objSrvHTTP.status <> 200 Then
	Response.Write "Link CIELO congestionado, tentando novamente em 5 segundos por favor aguarde"
	%>
	<script>setTimeout(window.location.href=window.location.href,10000)</script>
	<%
	Response.End
End If
   'resposta
   xml = objSrvHTTP.responseXML.xml

else

		Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

        xmlhttp.Open "POST",SoapServer,"false"
  		 xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
        xmlhttp.setRequestHeader "Content-Length", cStr(len(sRequest))

        xmlhttp.send "mensagem="&sRequest
        xml= xmlhttp.ResponseText 

   'resposta
   'xml = objSrvHTTP.responseXML.xml

end if   	


    retorno_codigo_erro = pegaValorNode(xml,"erro//codigo")
    retorno_mensagem_erro = pegaValorNode(xml,"erro//mensagem")
	
	if retorno_mensagem_erro <> "" then
		objConn.EXECUTE("INSERT INTO dadosCielo_movimento (processoId, meio, acao, resposta) VALUES ('"&request.Cookies("micCielo")("processoId")&"','1','Cancelamento','ERRO: "&REPLACE(retorno_mensagem_erro,"'"," ")&"')")
		response.write retorno_mensagem_erro
		response.End()
	end if

    retorno_tid = pegaValorNode(xml,"transacao//tid")



'cancelamento	Nó com dados do cancelamento caso tenha passado por esta etapa.
	cancelamentoCodigo = pegaValorNode(xml,"transacao//cancelamento/codigo")
	cancelamentoMsg = pegaValorNode(xml,"transacao//cancelamento/mensagem")
	cancelamentoDataHora = pegaValorNode(xml,"transacao//cancelamento/data-hora")
	cancelamentoValor = pegaValorNode(xml,"transacao//cancelamento/valor")	



		
	SQL = "UPDATE dados_cielo SET" & _
		  " cancelamentoCodigo = '"&cancelamentoCodigo&"', " & _
		  " cancelamentoMsg = '"&cancelamentoMsg&"', " & _
		  " cancelamentoDataHora = '"&cancelamentoDataHora&"', " & _
		  " cancelamentoValor = '"&cancelamentoValor&"', " & _
		  " cancelado = '1' " & _
		  
		  
		  "  WHERE pedidoId = '"&processo&"'"
		  
		  objConn.EXECUTE(SQL)
	
objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processoId&"','Autorização Cancelada')")
	

response.write "Cancelado Com sucesso."

%>