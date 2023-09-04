<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<!--#include file="../Library/Common/enviaEmail.asp" --> 
<!--#include file="../Library/Products/enviaGeral.asp" -->
<!--#include file="../Library/Products/funcaoLiquido.asp" -->
<!--#include file="../Library/Products/funcoesDig.asp" -->
<!--#include file="../Library/Products/PriceFunctions.asp" -->

<%
Session.LCID = 1046
response.Buffer = TRUE
Server.ScriptTimeout = 90000000

processo=request.QueryString("processo")


'if Request.cookies("FCNET_MIC")("idAge") = "" then
	set pegaClienteIdRs = objconn.execute("SELECT top 1 clienteId FROM vouchertemp WHERE processoId='"&processo&"' order by id")
	response.cookies("FCNET_MIC")("idAge") = pegaClienteIdRs("clienteId")
	idAge = pegaClienteIdRs("clienteId")
'else
'	idAge = Request.cookies("FCNET_MIC")("idAge")
'end if

if pegaClienteIdRs.eof then
	response.Write("Ocorreu um erro na emiss�o do processo")
	response.End()
end if

vAcao = Request.QueryString("AcaoSubmit")

if request.QueryString("tavola") = "S" then

	'hitorico do processo
	objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Pgto de boleto confirmado manualmente')")
	objConn.Execute("UPDATE  emissaoProcesso set pgtoAprovado='2', statusPGBoleto = '1', obs=obs+'|pagamento de boleto autorizado via Tavola' WHERE id='"&processo&"'")

end if



%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Affinity :: Assistencia de Viagem</title>
<link rel="stylesheet" href="../../../CSS/main.css">
</head>
<body leftmargin="10" topmargin="10" marginwidth="0" marginheight="0" >
<div id="divAguardando" align="center" style="background-color:#FFF; width:100%;" > 
  <p>&nbsp;</p>
  <p><br />
   <br />
   <%if vAcao = "0" Then 'se a opção for emitir voucher%>
  		Emiss&atilde;o em andamento.<br />
   <%else%>
   		<font size="1"><b>Gravando reserva<br />
   <%end if%>
Aguarde...</b></font></p>
<img src="../Images/loading.gif" width="220" height="19" />

<%
response.Flush()

codigoFree=Request.QUERYSTRING("codigoFree")

codigoFree  =REPLACE(REPLACE(TRIM(REPLACE(codigoFree,", ","")),"*",""),"-","")

if request.QueryString("pagamento") <> "" then
	pagamento=request.QueryString("pagamento")
else
	pagamento = request.QueryString("pg")
end if

ccMarca=request.QueryString("ccMarca")
tid=request.QueryString("tid")
autorizacao=request.QueryString("autorizacao")
emailTitular = request.QueryString("emailTitular")
sequenciaVouchers="0"
familiarSeg=1
ivoucher=0
antiLoopingInfinito=0
complementoCC="" 
 
vAcao = Request.QueryString("AcaoSubmit")

Set  cliRS = objConn.Execute("SELECT promotor1,promotor2,promotor3,  repId,ni from cadCliente where id='"&idAge&"'")

if pagamento = "AV" and request.QueryString("tavola") <> "S" then vAcao = 2 'for�a entrar na rotina de boleto



if request.QueryString("tavola") = "S" and request.QueryString("pg") = "AV" then pagamento = "AV" 'para libera��o de pagamento via t�vola para pgto AV

if request.QueryString("tavola") = "S" and pagamento = "" then
	set rsVTemp = objConn.execute("select top 1 pagamento from voucherTemp where processoid = '"&request.QueryString("processo")&"'")
	if not rsVTemp.eof then pagamento = rsVTemp("pagamento") 
end if

objconn.execute("update emissaoProcesso set pgtoForma= '"&pagamento&"', obs=obs+'|emitir.asp' where id='"&request.QueryString("processo")&"'")

dim ivoucher,codvoucher

set parcelaRs = objconn.execute("select parcelas,id,nPax,valorTotalBRL,emailRecibo,pgtoForma,pgtoAprovado,reservado,statusPGBoleto,boleto from emissaoProcesso where id='"&processo&"' ")

if request.QueryString("tavola") <> "S" then
	'--- bloqueia caso a forma de pagamento esteja diferente
	if parcelaRs("pgtoForma")<>ucase(pagamento) and vAcao <> 2 and vAcao <> 3 then
	response.Write "<script>alert('Ocorreu um erro no processo de emissao.\nPor favor, inicie o processo novamente.');"
	objconn.execute("update emissaoProcesso set obs=obs+'|bloqueado: forma de pagamento recebida("&pagamento&") n�o corresponde.' where id='"&processo&"'")
	response.Write "document.location='../home/planos.asp'</script>"
	response.End()
	end if
	
	'--- bloqueia caso o pagto CC nao esteja aprovado
	if parcelaRs("pgtoForma")="CC" and parcelaRs("pgtoAprovado")<>2 and vAcao <> 2 and vAcao <> 3 then
	response.Write "<script>alert('Ocorreu um erro no processo de emissao, confirma��o de pagamento n�o encontrada.\nPor favor, inicie o processo novamente.');"
	objconn.execute("update emissaoProcesso set obs=obs+'|bloqueado: status de pagamento ("&parcelaRs("pgtoAprovado")&") n�o aprovado.' where id='"&processo&"'")
	response.Write "document.location='../home/planos.asp'</script>"
	response.End()
	end if
end if
 
function geraCodVoucher(i,ageId,agencia,plano)
	antiLoopingInfinito=antiLoopingInfinito+1
	if antiLoopingInfinito > 109 then
		enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","filipe@ficopola.net","Affinity Assistencia Ocorreu um erro no emitir.asp (cod. do voucher nao foi gerado, [acionado pelo antiLoopingInfinito]) | processo: "&request.QueryString("processo"),"Ocorreu um erro. <br><br>valores das variaveis processo,pagamento,ccMarca,codVoucher,ivoucher:<br>"&processo&", "&pagamento&", "&ccMarca&", "&codVoucher&", "&ivoucher,0

		response.Write "Ocorreu um erro no processo de emiss�o."
		response.End()
	end if
	

	'set rsNplano = objConn.execute("select nPlano from planos where id  in (select planoId from emissaoProcesso where id = '"&request.QueryString("processo")&"')")

	
	'CRIA VOUCHER
	'verifica qual o proximo voucher para esta agencia
	if i=0 then
		'verifica qual o proximo voucher para esta agencia
		Set voucherAgeRS2 = objConn.Execute("SELECT top 1 voucherAge FROM voucher WHERE clienteid='"&ageId&"' order by voucherAge desc")
		if voucherAgeRS2.eof then
			 voucherAge = 100
		else
			 voucherAge =(voucherAgeRS2(0) + 1)
		end if
	else
		voucherAge=i+1
	end if

	if voucher1 = 0 then voucher1=voucherAge
 	

	'Atualiza o n. voucher para esta agencia
 	'objConn.Execute("UPDATE  ageUsers set nVouchers='"&voucherAge&"' WHERE id='"&ageId&"'")

	WHILE LEN(voucherAge) < 7
		voucherAge = "0" & voucherAge
	wend
	
	dayN = "day"&i
	dia = Request.Form(dayN)
	
	monthN = "month"&i
	mes = Request.Form(monthN)
	
	yearN = "year"&i
	ano = Request.Form(yearN)
	 
	dtNascimento = dia &"/"& mes &"/"& ano
	
	
	if request.QueryString("tavola") = "S" then prefixo = "T" else prefixo = "E"
	
	codvoucherf = prefixo & ucase(cliRs("ni")) &"/"& voucherAge & "-" & plano
	

	'if familiarSeg = "1" then 
'		codvoucherf = codvoucherf & ".F" &nPaxC
'	end if
	'FIM CRIA VOUCHER
	
	ivoucher = voucherAge
	voucherAgef=ivoucher

		
		set existeVoucherRs=objconn.execute("select id from voucher where voucher='"&codvoucherf&"'")
		if existeVoucherRs.EOF then
			if codvoucherf="" or ivoucher=0 then
				geraCodVoucher ivoucher,ageId,agencia,plano
			else
				'atualiza o numero de voucher para esta agencia
				objConn.Execute("UPDATE cadCliente SET nVouchers = '"&ivoucher&"' WHERE id='"&ageId&"'")
				codvoucher=codvoucherf
			end if
		else
			geraCodVoucher ivoucher,ageId,agencia,plano
		end if
end function


if request.QueryString("tavola") <> "S" then
'---PULA a trava para o login MODELO
'if lcase(Request.cookies("FCNET_MIC")("login"))<>"modelo" then
set verificaRs = objconn.execute("select nome,sobrenome from voucher where processoId='"&processo&"'")
if not verificaRs.EOF then
response.Write "<script>alert('J� existem vouchers emitidos com os dados fornecidos:"
while not verificaRs.EOF
response.Write "\n"&verificaRs(0)&" "&verificaRs(1)
verificaRs.movenext
wend
response.Write "');" 
objconn.execute("update emissaoProcesso set obs=obs+'|bloqueado:ja existe voucher emitido com este processoId' where id='"&request.QueryString("processo")&"'")
if pagamento="CC" then complementoCC="?atividade=1&processoId="&processo&"&aut=1"
response.Write "document.location='result.asp?voucher="&codVoucher&"&processo="&processo&"'</script>"
response.End()
end if
''end if 
'--------
end if

if vAcao = "0" Then 'se a op��o for emitir voucher

		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Inicio de gera��o de voucher')")



set paxTempRs = objconn.execute("select * from vouchertemp where processoId='"&processo&"' order by id")

while not paxTempRs.EOF
	clienteId=paxTempRs("clienteId")
	geraCodVoucher ivoucher,clienteId,paxTempRs("agencia"),paxTempRs("plano")
	if paxTempRs("familiar") = "1" and isnull(paxTempRs("flagUp")) then		
		codvoucher = codvoucher & ".F"&familiarSeg
		familiarSeg = familiarSeg + 1
	end if
		
	'if codvoucher="" or ivoucher=0 then
	'	geraCodVoucher ivoucher,clienteId,paxTempRs("agencia"),paxTempRs("plano")
	'end if

	if codvoucher="" or ivoucher=0 then
		enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","filipe@ficopola.net","Affinity Assistencia Ocorreu um erro no emitir.asp (cod. do voucher nao foi gerado) | processo: "&request.QueryString("processo"),"Ocorreu um erro. <br><br>valores das variaveis processo,pagamento,ccMarca,codVoucher,ivoucher:<br>"&processo&", "&pagamento&", "&ccMarca&", "&codVoucher&", "&ivoucher,0
		response.Write "Ocorreu um erro no processo de emiss�o."
		response.End()
	end if

	'-- Verifica se a ferramenta j� est� emitindo o processo
	set veriricaProcessoFerramentaRs = objconn.execute("select id from emissaoProcesso where obs like '%|emitido pela ferramenta do index.asp%' and id='"&processo&"'")
	if not veriricaProcessoFerramentaRs.EOF then
		enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","filipe@ficopola.net","Affinity Assistencia Processo de emissao interrompido | Numero: "&request.QueryString("processo"),"Detectado uma nova janela aberta com a ferramenta do index.asp emitindo o processo.<br>Notificacao automatica enviada pela ferramenta de bloqueio de duplicidade de voucher",0
		response.Write "<script>alert('Uma nova janela interrompeu o processo de emiss�o.');document.location='../home/default.asp'</script>"
		response.End() 
	end if
	'----------

	valorLiquido = calculaLiquido(paxTempRs("id"))
	
	nome_usuario = request.cookies("FCNET_MIC")("login")
	id_usuario = request.cookies("FCNET_MIC")("id")
	
	if nome_usuario = "" then nome_usuario = paxTempRs("emissorLogin")
	if id_usuario = "" then id_usuario = paxTempRs("emissor")
	
	if trim(nome_usuario) = "" or trim(id_usuario) = "" then
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Voucher sem autentica��o de emissor')")
	else
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Voucher com autentica��o de emissor')")
	end if
	
	objConn.execute("update emissaoProcesso set reservado='1', obs=obs+'| Salvando reserva' where id = '"&request.QueryString("processo")&"'")
	
	
	sequenciaVouchers=sequenciaVouchers&";"&codvoucher
	'Insere os dados na tabela voucher
	
	dataAV = date
	
	if pagamento = "AV" then
	dataEmite = dataAV
	else 
	dataEmite = paxTempRs("dataEmissao")
	end if

	premium = CalcPremium (paxTempRs("plano"), paxTempRs("dias"), paxTempRs("destino"), paxTempRs("idade"), codvoucher, paxTempRs("seqPro"))

	objConn.Execute("INSERT INTO voucher (voucher,seqPro,dataEmissao,horaEmissao,cambio,meio,agencia,voucherAge,representante,plano,inicioVigencia,fimVigencia,dias,totalBRL,totalUSD,comissaoUSD,comissaoBRL,netUSD,netBRL,overUSD, overBRL,destino,nome,sobrenome,beneficiario_nome, beneficiario_cpf, documento,email,celular,tipoDoc,endereco,fone,cidade,uf,cep, numero, bairro, complemento, idade, sexo, pais, familiar,emitido,emissor,emissorLogin, pagamento, inicioViagem, fimViagem, fileCli, acordo, planoAcordo,cancelado,clienteid,valorliquidoBRL,promotor1,promotor2,promotor3, comissaoBRL1,comissaoBRL2,dtNascimento, planoId, processoid, contatoNome, contatoFone, contatoEndereco, hotel, endHotel, foneHotel, nfCpf, liquidoBRL, bolTipo, bolCPF, bolNome, bolEnd, bolBairro, bolCep, bolCidade, bolUF, bolEmail, tipoGravidez,flagUp, semanas_gestacao, valor_premio_USD) VALUES ('"&codvoucher&"','"&paxTempRs("seqPro")&"','"&data(dataEmite,2,0)&"','"&paxTempRs("horaEmissao")&"','"&forMoeda(paxTempRs("cambio"),2)&"','"&paxTempRs("meio")&"','"&paxTempRs("agencia")&"','"&ivoucher&"','"&paxTempRs("representante")&"','"&paxTempRs("plano")&"','"&data(paxTempRs("inicioVigencia"),2,0)&"','"&data(paxTempRs("fimVigencia"),2,0)&"','"&paxTempRs("dias")&"','"&forMoeda(paxTempRs("totalBRL"),2)&"','"&forMoeda(paxTempRs("totalUSD"),2)&"','"&forMoeda(paxTempRs("comissaoUSD"),2)&"','"&forMoeda(paxTempRs("comissaoBRL"),2)&"','"&forMoeda(paxTempRs("netUSD"),2)&"','"&forMoeda(paxTempRs("netBRL"),2)&"','"&forMoeda(paxTempRs("overUSD"),2)&"','"&forMoeda(paxTempRs("overBRL"),2)&"','"&paxTempRs("destino")&"','"&paxTempRs("nome")&"','"&paxTempRs("sobrenome")&"','"&paxTempRs("beneficiario_nome")&"','"&paxTempRs("beneficiario_cpf")&"','"&paxTempRs("documento")&"','"&paxTempRs("email")&"','"&paxTempRs("celular")&"','"&paxTempRs("tipoDoc")&"','"&paxTempRs("endereco")&"','"&paxTempRs("fone")&"','"&paxTempRs("cidade")&"','"&paxTempRs("uf")&"','"&paxTempRs("cep")&"','"&paxTempRs("numero")&"','"&paxTempRs("bairro")&"','"&paxTempRs("complemento")&"','"&paxTempRs("idade")&"','"&paxTempRs("sexo")&"','"&paxTempRs("pais")&"','"&paxTempRs("familiar")&"','1','"&paxTempRs("emissor")&"','"&paxTempRs("emissorLogin")&"','"&pagamento&"','"&paxTempRs("inicioViagem")&"','"&paxTempRs("fimViagem")&"','"&paxTempRs("fileCli")&"','"&paxTempRs("acordo")&"','"&paxTempRs("planoAcordo")&"',0,'"&paxTempRs("clienteId")&"','"&forMoeda(paxTempRs("valorLiquidoBRL"),2)&"','"&paxTempRs("promotor1")&"','"&paxTempRs("promotor2")&"','"&paxTempRs("promotor3")&"','"&forMoeda(paxTempRs("comissaoBRL1"),2)&"','"&forMoeda(paxTempRs("comissaoBRL2"),2)&"','"&paxTempRs("dtNascimento")&"','"&paxTempRs("planoId")&"','"&processo&"','"&paxTempRs("contatoNome")&"','"&paxTempRs("contatoFone")&"','"&paxTempRs("contatoEndereco")&"','"&paxTempRs("hotel")&"','"&paxTempRs("endHotel")&"','"&paxTempRs("foneHotel")&"','"&paxTempRs("nfCpf")&"','"&forMoeda(valorLiquido,2)&"','"&paxTempRs("bolTipo")&"', '"&paxTempRs("bolCPF")&"','"&paxTempRs("bolNome")&"','"&paxTempRs("bolEnd")&"','"&paxTempRs("bolBairro")&"','"&paxTempRs("bolCep")&"','"&paxTempRs("bolCidade")&"','"&paxTempRs("bolUF")&"','"&paxTempRs("bolEmail")&"', '"&paxTempRs("tipoGravidez")&"', '"&paxTempRs("flagUp")&"', '"&paxTempRs("semanas_gestacao")&"', '"&forMoeda(premium,4)&"')")	
	
	objConn.execute("INSERT INTO voucherHistorico (voucher, usuario, obs) values ('"&codvoucher&"', '"&request.cookies("FCNET_MIC")("login")&"', 'Emiss�o pelo Site.<br>Navegador: "&request.servervariables("HTTP_USER_AGENT") &".<br>IP: "&Request.ServerVariables("REMOTE_ADDR")&"') ")
	
	objConn.Execute("UPDATE voucherTemp SET concluido = '1' WHERE id='"&paxTempRs("id")&"'")
	
	if pagamento="CC" then
	'REGISTRAR NA TABELA CC
		sql= "insert into cc (voucher,processoid,ccMarca,tid,autorizaDep,ccdataaut,ccNumero,ccTitular,parcelas) values ('"&codvoucher&"',"&processo&",'"&ccMarca&"','"&tid&"','"&autorizacao&"','"&data(date,2,0)&"','"&paxTempRs("numeroCartao")&"','"&paxTempRs("ccTitular")&"','"&parcelaRs(0)&"')"
		objConn.Execute(sql)
		aut="1"
	else
		aut="0"
	end if
	
	nPlano4Free = paxTempRs("plano")
	
	if pagamento = "FR" then objConn.Execute("UPDATE solicitaFree set voucherEmitido='"&codvoucher&"' WHERE voucherEmitido='' and codigoFree = '"&codigoFree&"' AND idAge = '"&Request.cookies("FCNET_MIC")("idAge")&"'")


	paxTempRs.movenext
wend

objconn.execute("update emissaoProcesso set usuarioLogin = '"&nome_usuario&"', obs=obs+'|while de insert de vouchers conclu�do' where id='"&request.QueryString("processo")&"'")
	
		
if ucase(pagamento) = "CC" then
'email de recibo
	voucherRecibo=split(sequenciaVouchers,";")
	
	HTML = "<table align=center width=500 border=0 cellpadding=0 cellspacing=0><tr><td align=left bgcolor=#FFFFFF><img src=http://www.affinityassistencia.com.br/imgs/logoAffinityP.png /></td></tr><tr><td bgcolor=#F2F6F7><table cellspacing=0 cellpadding=10><TR><TD width=100% valign=top borderColor=#c0c0c0><p><font size=1 face=verdana>Prezado Cliente,</font><br><br><font size=1 face=verdana> Agradecemos sua prefer&ecirc;ncia por nossa assist&ecirc;ncia em viagem.<br><br>Esta transa&ccedil;&atilde;o refere-se &agrave; compra dos seguintes vouchers de assist&ecirc;ncia em viagem Affinity Assistencia, feitas em nosso site www.affinityassistencia.com.br pela ag&ecirc;ncia de viagens <b>"&achaCliente(clienteId)&"</b>:"
		 for i=1 to ubound(voucherRecibo)
	HTML = HTML &"<br>"&voucherRecibo(i)
		 next
	HTML = HTML &"<br><br></font><font size=1 face=verdana>Seguem os dados da aprova&ccedil;&atilde;o&nbsp;de sua operadora de cart&atilde;o de  cr&eacute;dito:<BR>&nbsp;<BR><b>Processo:</b> 0000"&parcelaRs("id")&"<BR>    <b>N&uacute;mero de Passageiros:</b>"&parcelaRs("nPax")&" <BR>  <b>Valor Total:</b> R$ "&formatNumber(parcelaRs("valorTotalBRL"))&"<BR><b>N� de Parcelas:</b> "&parcelaRs("parcelas")&"    <BR>    </font><br>    <br>   <font size=1 face=verdana>Em sua fatura de cart�o de cr�dito dever� constar um pagamento para <B>Aspas Turismo</B>. Caso tenha d&uacute;vidas, entre em contato conosco atrav�s do   e-mail: <A href=mailto:suporte@affinityassistencia.com.br>suporte@affinityassistencia.com.br</A> informando o n&uacute;mero de seu processo: <b>0000"&parcelaRs("id")&"  </b>.</font></p></TD></TR><TR><TD borderColor=#c0c0c0 width=100% ><font color=#666666 size=1 face=verdana>Este &eacute; um e-mail autom&aacute;tico. N&atilde;o &eacute; necess&aacute;rio   respond&ecirc;-lo</font></TD>   </TR></table></td></tr><tr><td bgcolor=#DDE3E6><font size=1 face=verdana>&nbsp;&nbsp;&nbsp;&nbsp;Rio de Janeiro - Fone: +55 ( 21) 2531-1115</font></td></tr><tr>    <td height=15 bgcolor=#000549><font size=2><b><font color=#00CCFF><font color=#FFFFFF size=1 face=Arial, helvetica,, sans-serif>&nbsp;&nbsp;&nbsp;&nbsp;Affinity Assistencia "&YEAR(NOW)&"	</font></font></b></font></td></tr></table>"
	
	if parcelaRs("emailRecibo") <> "" and not isnull(parcelaRs("emailRecibo")) then
		enviaMail "no-reply@nextseguroviagem.com.br","Next Seguro Viagem","",parcelaRs("emailRecibo"),"Confirmacao de Compra",HTML,1

		'enviaMail "noreply@affinityassistencia.com.br","Affinity Assistencia","",parcelaRs("emailRecibo"),"Confirmacao de Compra",HTML,1
	end if
		'enviaMail "noreply@affinityassistencia.com.br","Affinity Assistencia","","suporte@affinityassistencia.com.br","Confirmacao de Compra",HTML,1
	'enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","filipe@ficopola.net","Confirma��o de Compra",HTML,1
	'enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","andrew.augusto@ficopola.net","Confirma��o de Compra",HTML,1

end if

'if request.cookies("FCNET_MIC")("login") = "modelo" then
'	enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","andrew.augusto@ficopola.net","Confirma��o de Compra",HTML,1
'	response.Write(HTML)
'	response.End()
'end if

objConn.Execute("update emissaoProcesso set finalizado='1',pgtoAprovado='2' where id='"&processo&"'")

		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Fim de gera��o de voucher')")


'ENVIA EVOUCHER POR EMAIL E SMS
enviaEvoucher(processo) 
 


if err.description <> "" then
enviaMail "web@affinityassistencia.com.br","Affinity Assistencia","","filipe@ficopola.net","Erro na pagina de emissao - emitir.asp | processo: "&request.QueryString("processo"),"Ocorreu um erro <br><br>" & Err.number & " - " & err.Description 	& "<br> Contexto="& err.helpcontext &"<br> Origem="& err.nativeerror &"<br>  Fonte="& err.source,0
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Ocorreu um erro')")
End If


end if 'Fim op��o EMITIR VOUCHER

if vAcao = "0" Then 'reserva confirmada, status 2
	objConn.execute("update emissaoProcesso set reservado='2' where id = '"&request.QueryString("processo")&"'")
	if request.QueryString("tavola") = "S" then 'ap�s autorizado o pagamento do boleto, volta para o t�vola%>
		<script>document.location='http://tavola.ficopola.net:8099/processo/detalhes.asp?processo=<%=request.QueryString("processo")%>'</script>
	<%end if
end if

if vAcao = "1" Then 'Op��o Salvar e Continuar
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Processo Gravado')")

objConn.execute("update emissaoProcesso set reservado='1', obs=obs+'| Salvando reserva' where id = '"&request.QueryString("processo")&"'")
%>
	<script>document.location='formaPagamento2.asp'</script>
<%end if

if vAcao = "2"  Then 'Op��o Salvar e Sair / Pagamento via boleto
	
	if pagamento = "AV" then 's� salvando
		if request.QueryString("AcaoSubmitBoleto") = 0 then
			
			objConn.execute("update emissaoProcesso set reservado='1', obs=obs+'| Salvando reserva de pagamento via boleto' where id = '"&request.QueryString("processo")&"'")
			%>
				<script>document.location='../emitidos/indexBoleto.asp'</script>
			<%
		else
			objConn.execute("update emissaoProcesso set reservado='1', obs=obs+'| Gerando boleto' where id = '"&request.QueryString("processo")&"'")
			set tipoBoletoRS = objConn.execute("SELECT coalesce(bolTipo,0) as bolTipo from voucherTemp where processoID = '"&request.QueryString("processo")&"'")
			call gerarBoleto(request.QueryString("processo"),tipoBoletoRS("bolTipo"))
			if tipoBoletoRS("bolTipo") = "0" then boletoParaSTR = "para agencia" else boletoParaSTR = "para PAX"
			
			'hitorico do processo
			objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&request.QueryString("processo")&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Boleto Gerado "&boletoParaSTR&"')")

			if tipoBoletoRS("bolTipo") = "1" then ' boleto para pax
			
				
				'''''''''''''''''''''' EMAIL INICIO
				Set objDynu = Server.Createobject("Dynu.HTTP") 
				objDynu.SetURL "http://www.affinityassistencia.com.br/bibliotecaDocs/notificacao.asp?processo="&request.QueryString("processo")
				HTML = objDynu.PostURL()
				Set objDynu = nothing 
				set voucher_TempRs = objconn.execute("select * from voucherTemp where processoId='"&request.QueryString("processo")&"' order by id")
				
				if not isnull(voucher_TempRs("email")) and trim(voucher_TempRs("email")) <> "" then
					enviaMail "no-reply@affinityassistencia.com.br","Affinity Assistencia","",voucher_TempRs("email"),"Affinity Assistencia - Boleto",HTML,1
			'hitorico do processo
			objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&request.QueryString("processo")&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Boleto Envaido para "&voucher_TempRs("email")&"')")
				end if
			
			
			end if
			
			'''''''''''''''''''''' EMAIL FIM
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Direcionado para Boleto')")
			%>
				<script>document.location='../emitidos/indexBoleto.asp'</script>
			<%
		end if
		
	else 'processos que n�o sao pagamento AV ou que n�o foi informado o pagamento
	
		objConn.execute("update emissaoProcesso set reservado='1', obs=obs+'| Salvando reserva' where id = '"&request.QueryString("processo")&"'")
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Direcionado para reservas')")
		%>
			<script>document.location='../emitidos/indexReserva.asp'</script>
		<%
	end if
end if

if vAcao = "3" Then 'Op��o Salvar e Sair
	objConn.execute("update emissaoProcesso set reservado='1', obs=obs+'| Salvando reserva' where id = '"&request.QueryString("processo")&"'")
	%>
		<script>document.location='../emitidos/indexCadCartaoGarantia.asp?processo=<%response.Write(request.QueryString("processo"))%>'</script>
<%end if



if request.QueryString("ccMarca")="VI" then 
%>
	<script> 
    document.location='result.asp?voucher=<%=codVoucher%>&processo=<%=processo%>'
    </script>
    
    <!--script>
    document.getElementById('divAguardando').style.display='none'
    window.open('result.asp?voucher=<%=codVoucher%>','principal')
    </script --><small><small>
    Emiss�o conclu�da com sucesso.<br><a href="result.asp?voucher=<%=codVoucher%>&processo=<%=processo%>" target="_blank"> Clique aqui</a> caso a p�gina n�o tenha sido redirecionada automaticamente.</small></small>
<%
'Fecha o Objeto de Conexão
objConn.close
Set objConn = Nothing 

else
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Direcionado para listagem de voucher')")

%>
	<script>
    document.location='result.asp?voucher=<%=codVoucher%>&processo=<%=processo%>'
    </script>
<%
'Fecha o Objeto de Conexão
objConn.close
Set objConn = Nothing 

end if%>
</div>

