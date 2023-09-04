<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<!--#include file="../Library/Products/Cupom.asp" -->

<%
	'processo
	processoId = request.form("processo")
	'plano
	planoId = protetorSQL(request.form("plano"))
	'usuario
	usuario = Request.cookies("wlabel")("siteNome")
	'agencia
	agencia = Request.cookies("wlabel")("revId")
	'cupom
	cupom = protetorSQL(request.form("cupom"))

	brl = request.form("brl")

	usd = request.form("usd")

	totalBRL = request.form("totalBRL")
	
	
	if processoId = "" or cupom = "" or planoId = "" then
		response.write "Preencha um cupom válido"
		response.end()
	end if
	
	
	aplicaCupom cupom, processoId, agencia, planoId, brl, usd, totalBRL
	
	response.end()
	
%>

