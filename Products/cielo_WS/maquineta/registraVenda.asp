<!--#include file="../../../Library/Common/micMainCon.asp" -->
<!--#include file="../../../Library/Common/funcoes.asp" -->
<!--#include file="websevice.asp" -->
<%
Response.AddHeader "PRAGMA", "NO-CACHE" 
response.expires=0

	processo = request.form("orderid")
	parcelas = request.form("parcelas")
	
	numCartao = TRIM(REPLACE(request.form("number")," ",""))
	validadeCartao = right(request.form("expiry"),4)&left(request.form("expiry"),2)
	codSeg = request.form("cvc")
	ccMarca = request.form("ccMarca")
	Select case ccMarca 
		Case "visa"
			ccMarca = "VI"
  		Case "mastercard"
		  	ccMarca = "MA"
		Case "dinersclub"
			ccMarca = "DI"
		Case "amex"
			ccMarca = "AM"
		Case "elo"
			ccMarca = "EL"
  		Case else
		  	ccMarca = "TESTE"
		End Select
	V_NGrava = TRIM(REPLACE(request.form("number")," ",""))
	titular = request.form("name")
	
	''''''''''''''''''''''''' verificar se o processo ja possui tid aprovado
	set testeAprovadoRS = objConn.execute("SELECT * FROM dados_cielo where pedidoId='"&processo&"' and (statusCielo='4' or statusCielo='6') ")
	if not testeAprovadoRS.EOF then
	tid 		= testeAprovadoRS("tid")
	autorizacao = testeAprovadoRS("autorizacaoArp")
	testeAprovadoRS.CLOSE
	set testeAprovadoRS = nothing
	
	objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processo&"','Processo ja aprovado TID: "&tid&" COD: "&autorizacao&" | redirecionado para finalizar')")
	
	response.Redirect "autorizada.asp?processo="&processo&"&pagamento=CC&ccMarca=VI&tid="&tid&"&autorizacao="&autorizacao

	response.Redirect 
	response.End()
	end if
	testeAprovadoRS.CLOSE
	set testeAprovadoRS = nothing
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	
	V_NGrava = LEFT(V_NGrava,4) & "********" & RIGHT(V_NGrava,4)
		
    objConn.execute("INSERT INTO cielo_controle (processoId,numero,expira,codigo,emissor,parcelas,titular) VALUES ('"&processo&"','"&V_NGrava&"','"&validadeCartao&"','***','"&ccMarca&"','"&parcelas&"','"&titular&"')")
	objConn.execute("UPDATE emissaoProcesso set parcelas = '"&parcelas&"', ccMarca = '"&ccMarca&"' WHERE id = "&processo)
	
	objConn.EXECUTE("INSERT INTO dados_cielo (regIP, pedidoId) values ('"&Request.ServerVariables("REMOTE_ADDR")&"','"&processo&"')")


response.Cookies("Cielo")("processoId") = processo


	bandeira = lcase(ccMarca)
	parcelas = parcelas
	
	if bandeira = "vi" then bandeira = "visa"
	if bandeira = "ma" then bandeira = "mastercard"
	if bandeira = "di" then bandeira = "diners"
	if bandeira = "am" then bandeira = "amex"
	if bandeira = "el" then bandeira = "elo"
	
	Set rs= objConn.Execute("select coalesce(SUM(totalBRL),1) from vouchertemp where processoId="&processo)
	if rs.eof then response.write "Erro carregando dados do database processo total." : response.end
	
	if parcelas = "1" then
		produto = "1"
	else
		produto = "2"
	end if


	
	vlTotal= replace(replace(formatNumber(rs(0),2),",",""),".","")
	'vlTotal = 100
	
'if b2w = "S" then
'vlTotal = 100
'end if


	
	numeroPedido = processo
	valor = vlTotal
	datahoraT = YEAR(now) & "-" & RIGHT("0"&MONTH(now),2) & "-"  & RIGHT("0"&DAY(now),2) & "T" & TIME
	
	
sRequest = "<?xml version='1.0' encoding='ISO-8859-1' ?> " & vbcrlf & _
"<requisicao-transacao id='a97ab62a-7956-41ea-b03f-c2e9f612c293' versao='1.2.1'>"& vbcrlf & _
"	<dados-ec>" & vbcrlf & _
"		<numero>"&afiliacaoCielo&"</numero>" & vbcrlf & _
"		<chave>"&chaveWS&"</chave>" & vbcrlf & _
"	</dados-ec>" & vbcrlf & _
"	<dados-portador>" & vbcrlf & _
"		<numero>"&numCartao&"</numero>" & vbcrlf & _
"		<validade>"&validadeCartao&"</validade>" & vbcrlf & _
"		<indicador>1</indicador>" & vbcrlf & _
"		<codigo-seguranca>"&codSeg&"</codigo-seguranca>" & vbcrlf & _
"		<nome-portador>"&codSeg&"</nome-portador>" & vbcrlf & _
"		<token></token>" & vbcrlf & _
"	</dados-portador> " & vbcrlf & _
"	<dados-pedido>" & vbcrlf & _
"		<numero>"&numeroPedido&"</numero>"& vbcrlf & _
"		<valor>"&valor&"</valor>" & vbcrlf & _
"		<moeda>986</moeda>" & vbcrlf & _
"		<data-hora>"&datahoraT&"</data-hora>" & vbcrlf & _
"		<descricao>Emissao de voucher Affinity Assistencia "&processo&"</descricao>" & vbcrlf & _
"		<idioma>PT</idioma>" & vbcrlf & _
"		<soft-descriptor>Assist Viagem</soft-descriptor>" & vbcrlf & _
"	</dados-pedido>" & vbcrlf & _
"	<forma-pagamento>" & vbcrlf & _
"		<bandeira>"&bandeira&"</bandeira>" & vbcrlf & _
"		<produto>"&produto&"</produto>" & vbcrlf & _
"		<parcelas>"&parcelas&"</parcelas>" & vbcrlf & _
"	</forma-pagamento>"& vbcrlf & _
"	<url-retorno>http://www.affinityseguro.com.br/v20/teste/Products/cielo_WS/maquineta/retorno.asp?pedidoId="&processo&"</url-retorno>" & vbcrlf & _
"	<autorizar>3</autorizar>" & vbcrlf & _
"	<capturar>true</capturar>" & vbcrlf & _
"	<campo-livre>"&processo&"</campo-livre>" & vbcrlf & _
"	<bin>"&LEFT(numCartao,6)&"</bin>" & vbcrlf & _
"</requisicao-transacao>"

'if b2w = "S" then
'response.write sRequest
'response.End()
'end if
   

Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
xmlhttp.SetTimeouts 30000, 30000, 70000, 70000

        xmlhttp.Open "POST",SoapServer,"false"
  		 xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded charset=utf-8" 
        xmlhttp.setRequestHeader "Content-Length", cStr(len(sRequest))

        xmlhttp.send "mensagem="&sRequest
        xml= xmlhttp.ResponseText 

   'resposta
   'xml = objSrvHTTP.responseXML.xml
   
'if b2w = "S" then
'response.write xml
'response.End()
'end if



    retorno_codigo_erro = pegaValorNode(xml,"erro//codigo")
    retorno_mensagem_erro = pegaValorNode(xml,"erro//mensagem")
	
	if retorno_mensagem_erro <> "" then
	'objConn.EXECUTE("INSERT INTO dadosCielo_movimento (processoId, meio, acao, resposta) VALUES ('"&response.Cookies("Cielo")("processoId")&"','1','Postagem de registro.','"&retorno_mensagem_erro&"')")
	response.write retorno_mensagem_erro
	response.End()
	end if

    retorno_tid = pegaValorNode(xml,"transacao//tid")

	objConn.EXECUTE("INSERT INTO dadosCielo_movimento (processoId, meio, acao, resposta) VALUES ('"&request.Cookies("Cielo")("processoId")&"','1','Postagem de registro.','TID Retornado: "&retorno_tid&"')")


    retorno_pedido = pegaValorNode(xml,"transacao//dados-pedido//numero")
    retorno_valor = pegaValorNode(xml,"transacao//dados-pedido//valor")
    retorno_moeda = pegaValorNode(xml,"transacao//dados-pedido//moeda")
    retorno_data_hora = pegaValorNode(xml,"transacao//dados-pedido//data-hora")
    retorno_descricao = pegaValorNode(xml,"transacao//dados-pedido//descricao")
    retorno_idioma = pegaValorNode(xml,"transacao//dados-pedido//idioma")

    retorno_produto = pegaValorNode(xml,"transacao//forma-pagamento//produto")
    retorno_parcelas = pegaValorNode(xml,"transacao//forma-pagamento//parcelas")

    retorno_status = pegaValorNode(xml,"transacao//status")

    retorno_url_autenticacao = pegaValorNode(xml,"transacao//url-autenticacao")

    ' Se não ocorreu erro exibe parâmetros
    If retorno_codigo_erro = "" Then
       ' Response.write "<b> TRANSAÇÃO (ambiente fipola/mic brasil)</b><br>"
       ' Response.write "<b>Código de identificação do pedido (TID): </b>" & retorno_tid & "<br>"
       ' Response.write "<b>Número do pedido (numero): </b>" & retorno_pedido & "<br>"
       ' Response.write "<b>Valor do pedido (valor): </b>" & retorno_valor & "<br>"
       ' Response.write "<b>Moeda do pedido (moeda): </b>" & retorno_moeda & "<br>"
       ' Response.write "<b>Data e hora do pedido (data-hora): </b>" & retorno_data_hora & "<br>"
       ' Response.write "<b>Descrição do pedido (descricao): </b>" & retorno_descricao & "<br>"
       ' Response.write "<b>Idioma do pedido (idioma): </b>" & retorno_idioma & "<br>"
       ' Response.write "<b>Forma de pagamento (produto): </b>" & retorno_produto & "<br>"
       ' Response.write "<b>Número de parcelas (parcelas): </b>" & retorno_parcelas & "<br>"
       ' Response.write "<b>Status do pedido (status): </b>" & retorno_status & "<br>"
       ' Response.write "<b>URL para autenticação (url-autenticacao): </b>" & retorno_url_autenticacao & "<br>"
	  objconn.execute("update emissaoProcesso set obs=obs+'|WS CIELO retornou TID "&retorno_tid&"' where id='"&processo&"'")
		
	SQL = "UPDATE dados_cielo SET" & _
		  " tid = '"&retorno_tid&"', " & _
		  " pedidoValor = '"&retorno_valor&"', " & _
		  " pedidoMoeda = '"&retorno_moeda&"', " & _
		  " pedidoDataHora = '"&retorno_data_hora&"', " & _
		  " idioma = '"&retorno_idioma&"', " & _
		  " produto = '"&retorno_produto&"', " & _
		  " pedidoDescricao = '"&retorno_descricao&"', " & _
		  " pagamentoForma = '"&retorno_produto&"', " & _
		  " pagamentoParcelas = '"&retorno_parcelas&"', " & _
		  " statusCielo = '"&retorno_status&"' " & _
		  "  WHERE pedidoId = '"&processo&"'"
		  
		 ' Response.Write(SQL)
		  'Response.End()
		  objConn.EXECUTE(SQL)

    Else
        Response.write "<b>Erro: </b>" & retorno_codigo_erro & "<br>"
        Response.write "<b>Mensagem: </b>" & retorno_mensagem_erro & "<br>"
	  objconn.execute("update emissaoProcesso set obs=obs+'|ERRO WS CIELO: "&retorno_mensagem_erro&"' where id='"&processo&"'")
		response.End()
    End If
	
	  objconn.execute("update emissaoProcesso set obs=obs+'|WS CIELO direcionou para autenticação.' where id='"&processo&"'")
	  
	  
response.redirect "retorno.asp?pedidoId="&processo
%>