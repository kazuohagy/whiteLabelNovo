<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<%
jaExiste=0

processo = request.QueryString("processo")

'hitorico do processo
objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Inicio verificacao cartao')")


if  request("processo")="" then	response.write "ERRO: Nao foi recebido os parametro Processo" : response.End()

'salvando o email digitado, pois nunca era enviado email de compra CC
if request.QueryString("emailTitular") <> "" then
	objconn.execute("update emissaoProcesso set emailRecibo = '"&request.QueryString("emailTitular")&"' where id='"&request.QueryString("processo")&"'")
end if

labelCcMarca = ucase(left(request.QueryString("ccMarca"),2))


	' verificar se j� h� ovucher neste processo
	set verificaRs = objconn.execute("select nome,sobrenome from voucher where processoId='"&request("processo")&"'")
	if not verificaRs.EOF then
		response.Write "<script>alert('Já existem vouchers emitidos com os dados fornecidos"
		while not verificaRs.EOF
			response.Write "\n"&verificaRs(0)&" "&verificaRs(1)
			verificaRs.movenext
		wend
		
		
		objconn.execute("update emissaoProcesso set obs=obs+'|bloqueado:ja existe voucher emitido neste processo' where id='"&request.QueryString("processo")&"'")
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','bloqueado:ja existe voucher emitido neste processo')")
		response.Write "');"
		response.Write "document.location='../Account/Sales_History.asp'</script>"
		response.End()
	end if
	
	
	
	' verifica se ja tem aprovasao neste processo
	set verifica2Rs = objconn.execute("select id from dados_cielo where pedidoId='"&processo&"' and autorizado = 'S' and tid is not null")

	if not verifica2Rs.EOF then jaExiste=1 ' existe

	if jaExiste=1 then ' sim existe
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','bloqueado:ja existe TID aprovado para este pedido')")
		response.Write "<script>alert('Ja existe uma transacao de cartao de credito APROVADA para este processo de emissao.');"
		objconn.execute("update emissaoProcesso set obs=obs+'|bloqueado:ja existe dados de transacao neste processo' where id='"&request.QueryString("processo")&"'")
		response.Write "document.location='../Account/Sales_History.asp'</script>"
		response.End()
	end if

	  orderid = processo
	


	  ' nova logica cielo
	  'objConn.EXECUTE("INSERT INTO dados_cielo (regIP, pedidoId) values ('"&Request.ServerVariables("REMOTE_ADDR")&"','"&orderid&"')")
	  'objconn.execute("update emissaoProcesso set obs=obs+'|Redirecionado para consumo do web service CIELO' where id='"&request.QueryString("processo")&"'")
	  'hitorico do processo
	 ' objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Redirecionado para consumo do web service CIELO')")
	  
	  objConn.close : set objConn=nothing
	
	  response.Cookies("micCielo")("processoId") = processo
	  response.Redirect "formulario-seguro.asp?parcelas="&parcelas&"&orderid="&orderid&"&idAge="&request.cookies("wlabel")("revId")
%>