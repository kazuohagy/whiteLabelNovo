<!--#include file="../../Library/Common/micMainCon.asp" -->
<!--#include file="aspjson/aspJSON1.19.asp" -->
<!--#include file="apiCielo.asp" -->
<%
	processo = request.form("orderid")
	parcelas = request.form("parcelas")
	numCartao = TRIM(REPLACE(request.form("number")," ",""))
	validadeCartao = left(request.form("expiry"),2) & "/" & right(request.form("expiry"),4)
	codSeg = request.form("cvc")
	
	'padrao de bandeiras para cielo:
	'(Visa / Master / Amex / Elo / Aura / JCB / Diners / Discover / Hipercard / Hiper).
	ccMarca = request.form("ccMarca")
	'24/09/2021 apenas master e diners precisaram de tratamento
	if ccMarca = "mastercard" OR ccMarca = "dinersclub" then
		ccMarca = Left(ccMarca, 6)
	end if
	
	V_NGrava = TRIM(REPLACE(request.form("number")," ",""))
	titular = request.form("name")
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	V_NGrava = LEFT(V_NGrava,4) & "********" & RIGHT(V_NGrava,4)
		
    objConn.execute("INSERT INTO cielo_controle (processoId,numero,expira,codigo,emissor,parcelas,titular) VALUES ('"&processo&"','"&V_NGrava&"','"&validadeCartao&"','***','"&ccMarca&"','"&parcelas&"','"&titular&"')")
	
	objConn.execute("UPDATE emissaoProcesso set parcelas = '"&parcelas&"', ccMarca = '"&ccMarca&"' WHERE id = "&processo)
	
	objConn.EXECUTE("INSERT INTO dados_cielo (regIP, pedidoId) values ('"&Request.ServerVariables("REMOTE_ADDR")&"','"&processo&"')")

	response.Cookies("Cielo")("processoId") = processo

	bandeira = ccMarca
	parcelas = parcelas
	
	Set rs= objConn.Execute("select coalesce(SUM(totalBRL),1) from vouchertemp where processoId="&processo)
	if rs.eof then response.write "Erro carregando dados do database processo total." : response.end
	
	'valor da compra em centavos
	vlTotal= replace(replace(formatNumber(rs(0),2),",",""),".","")
	
	numeroPedido = processo
	valor = vlTotal
	datahoraT = YEAR(now) & "-" & RIGHT("0"&MONTH(now),2) & "-"  & RIGHT("0"&DAY(now),2) & "T" & TIME
	
	Set oJSON = New aspJSON

	With oJSON.data
		.Add "MerchantOrderId", processo

		.Add "Payment", oJSON.Collection()
		With oJSON.data("Payment")
			
			.Add "Type" , "CreditCard"
			.Add "Amount" , valor
			.Add "Installments" , parcelas
			.Add "Capture", "true"
			
			.Add "CreditCard", oJSON.Collection()
			With .item("CreditCard")
			
				.Add "CardNumber", numCartao
				.Add "Holder", titular
				'verificar formato
				.Add "ExpirationDate", validadeCartao 
				.Add "SecurityCode", codSeg
				.Add "Brand", bandeira
				
			End With
			
			.Add "IsCryptoCurrencyNegotiation", false
		
		End With

	End With
	
	'endpoint da venda simples com cartão de credito
	EndPoint = urlProducao & "/1/sales/"
	'envia request
	Set jsonHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	jsonHttp.SetTimeouts 30000, 30000, 70000, 70000
    jsonHttp.Open "POST", EndPoint, "false"
	jsonHttp.setRequestHeader "Content-Type", "application/json" 
    jsonHttp.setRequestHeader  "MerchantId", MerchantId
	jsonHttp.setRequestHeader  "MerchantKey", MerchantKey
    jsonHttp.send oJSON.JSONoutput() 
		
	json = jsonHttp.ResponseText 
	
	'nova instancia para o retorno
	Set aJSON = New aspJSON
	
	aJSON.loadJSON(json)
	
	'problema na comunicação com a cielo
	retorno_codigo_erro = aJSON.data("Code")
	retorno_mensagem_erro = aJSON.data("Message")
	if retorno_mensagem_erro <> "" OR  retorno_codigo_erro <> "" then
		objConn.EXECUTE("INSERT INTO dadosCielo_movimento (processoId, meio, acao, resposta) VALUES ('"&response.Cookies("Cielo")("processoId")&"','1','Postagem de registro.','"&retorno_mensagem_erro&"')")
		objconn.execute("update emissaoProcesso set obs=obs+'|ERRO API CIELO: "&retorno_mensagem_erro&"' where id='"&processo&"'")
		response.write "Ocorreu um erro inesperado. Tente novamente mais tarde."
		response.End()
	end if
	
	retorno_tid = aJSON.data("Payment").item("Tid")
    retorno_valor = aJSON.data("Payment").item("Amount")
    retorno_moeda = aJSON.data("Payment").item("Currency")
    retorno_data_hora = aJSON.data("Payment").item("ReceivedDate")
    retorno_descricao = aJSON.data("Payment").item("ReturnMessage")
    retorno_parcelas = aJSON.data("Payment").item("Installments")
    retorno_status = aJSON.data("Payment").item("Status")
	retorno_codigo = aJSON.data("Payment").item("ReturnCode")
	'utilizado para ler o retorno na url de consulta 
	id_pagamento = aJSON.data("Payment").item("PaymentId")
	
	objconn.execute("update emissaoProcesso set obs=obs+'|API CIELO retornou TID "&retorno_tid&"' where id='"&processo&"'")

	SQL = "UPDATE dados_cielo SET" & _
		  " tid = '"&retorno_tid&"', " & _
		  " pedidoValor = '"&retorno_valor&"', " & _
		  " pedidoMoeda = '"&retorno_moeda&"', " & _
		  " pedidoDataHora = '"&retorno_data_hora&"', " & _
		  " pedidoDescricao = '"&retorno_descricao&"', " & _
		  " pagamentoParcelas = '"&retorno_parcelas&"', " & _
		  " statusCielo = '"&retorno_status&"', " & _
		  " lr = '"&retorno_codigo&"', " & _
		  " urlAutenticacao = '"&id_pagamento&"' " & _
		  "  WHERE pedidoId = '"&processo&"'"
		  
	 
	objConn.EXECUTE(SQL)
	
	response.redirect "retorno.asp?pedidoId="&processo

%>