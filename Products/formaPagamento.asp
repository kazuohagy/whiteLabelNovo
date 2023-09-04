<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<!--#include file="../Library/Common/cambio.asp" -->
<!--#include file="../Library/Common/enviaEmail.asp" -->
<!--#include file="../Library/Products/PriceFunctions.asp" -->
<!--#include file="../Library/Products/Cupom.asp" -->
<%
Session.LCID = 1046

if 1 <> 1 then
	if request.cookies("FCNET_MIC")("fluxoEmissao")<>2  then response.Write "<script>alert('Prezado usuário, por motivo de segurança de dados não é possível recarregar o mesmo processo de venda.\nFavor iniciar o processo novamente.');document.location='../home/planos.asp'</script>" : response.End() 
end if	
	
response.cookies("FCNET_MIC")("fluxoEmissao")=3
familiarSeg = 1
processo = protetorSQL(Request.Form("processo"))
idPlano = protetorSQL(Request.Form("idPlano"))
idAge = protetorSQL(Request.cookies("wlabel")("revId"))
planoId = protetorSQL(Request.Form("planoId"))
parcelas= protetorSQL(request.form("parcelas"))
plano = protetorSQL(Request.Form("planoN"))
acordo = protetorSQL(Request.Form("acordo"))
ageNi = protetorSQL(Request.Form("ageNi"))
dias = protetorSQL(Request.Form("dias"))
destino = protetorSQL(Request.Form("destino"))
paxTotal = protetorSQL(Request.Form("paxTotal"))
familiar = protetorSQL(Request.Form("familiar"))
total =  protetorSQL(Request.Form("total"))
origemRecep = protetorSQL(request.Form("origemRecep")) 
pagamento = "CC"

if idAge = "28" then

	'pagamento = "FA"

end if 

'verifica se foi aplicado o cupom no processo
cupomDescontoId = ""

set vrfcCupomRs = objConn.execute("SELECT desconto FROM emissaoProcesso WHERE id = '"& processo &"'")

if not vrfcCupomRs.eof and vrfcCupomRs(0) <> "0" and vrfcCupomRs(0) <> "" and isNumeric(vrfcCupomRs(0)) then
	cupomDescontoId = vrfcCupomRs(0)
end if

objConn.Execute("UPDATE emissaoProcesso set desconto = '"&desconto&"',pgtoAprovado='4' WHERE id ='"&processo&"'")

if acordo = "1" then planoAcordo = protetorSQL(Request.Form("idPlano"))

inicioViagem2 =  protetorSQL(Request.Form("inicioViagem2"))
fimViagem2 =  protetorSQL(Request.Form("fimViagem2"))

inicioViagem =  protetorSQL(Request.Form("inicioViagem"))
fimViagem =  protetorSQL(Request.Form("fimViagem"))

cepPax			= protetorSQL(Request.Form("cepPax"))
enderecoPax		= replace(protetorSQL(Request.Form("enderecoPax")),"'"," ")
complementoPax	= replace(protetorSQL(Request.Form("complementoPax")),"'"," ")
bairroPax		= replace(protetorSQL(Request.Form("bairroPax")),"'"," ")
numeroPax		= replace(protetorSQL(Request.Form("numeroPax")),"'"," ")
bairroPax		=  replace(protetorSQL(Request.Form("bairroPax")),"'"," ")
dddFo			= protetorSQL(Request.Form("dddFo"))
foneN			= protetorSQL(Request.Form("foneN"))
fone =  foneN
cidadePax = replace(protetorSQL(Request.Form("cidadePax")),"'"," ")
ufPax = protetorSQL(Request.Form("ufPax"))
pais = protetorSQL(Request.Form("pais"))

contatoNome = replace(protetorSQL(Request.form("contatoNome")),"'"," ")
contatoDDD = protetorSQL(Request.form("contatoDDD"))
contatoFoneN = protetorSQL(Request.form("contatoFoneN")) 
contatoFone = contatoFoneN
contatoEndereco = replace(protetorSQL(Request.form("contatoEndereco")),"'"," ") 

cambio= protetorSQL(Request.Form("cambio"))

dataEmissao = data(date(),1,0)

horaEmissao = TIME()
meio = 2
representante = "0"
emissor = Request.cookies("wlabel")("xml_Id")
emissorLogin = Request.cookies("wlabel")("xml_A")

sobrePax1 = replace(protetorSQL(Request.Form("sobrePax1")),"'"," ")
nomePax1 = replace(protetorSQL(Request.Form("nomePax1")),"'"," ")

processo= request("processo")

vet_voucherTemp = split(protetorSQL(request.Form("idVoucherTemp")),",")

if acordo <> "1" then
	'Seleciona comissao do plano	
	Set rsCom = objConn.Execute("SELECT * FROM comissao where idAge ='"&idAge&"' AND  idPlano='"&idPlano&"'")
	if rsCom.EOF then
		rsCom.CLOSE
		set rsCom = nothing
		Set planoNovoRS = objConn.Execute("SELECT * FROM planos where id='"&idPlano&"'")
		Set planoAntigoRS = objConn.Execute("SELECT * FROM planos where publicado='1' and ageId='0' and versaoTarifa='4' and nPlano='"&planoNovoRS("nPlano")&"'")
		
		if planoAntigoRS.eof then 
			response.Write("Não há comissão registrada para este plano, por favor, entre em contato com o Affinity")
			response.End()
		end if
		Set rsCom = objConn.Execute("SELECT * FROM comissao where idAge ='"&idAge&"' AND  idPlano='"&planoAntigoRS("id")&"'")

		planoNovoRS.CLOSE
		set planoNovoRS = nothing
		planoAntigoRS.CLOSE
		set planoAntigoRS = nothing
	end if
end if

Set  cliRS = objConn.Execute("SELECT promotor1, promotor2, promotor3, ni, pgtoFA, pgtoCC, pgtoBO FROM cadCliente WHERE id='"&request.cookies("wlabel")("revId")&"'")
Set  userRS = objConn.Execute("SELECT comissao FROM usuarios WHERE id='"&cliRS("promotor1")&"'")
set regraTarifaRS = objconn.execute("select regraTarifa,nacionalidade, idadeMinima from planos where id='"&planoId&"'")

For pax_count=1 to paxTotal
	cambio = Request.Form("cambio"&pax_count)
	totalUSD = Request.Form("valorUSD_Origin"&pax_count)
	totalBRL = Request.Form("valorBRL_Origin"&pax_count)

	if destino = 4 then
		totalBRL = 0
	end if
	
	totalUSD = replace(totalUSD,".",",")
	totalBRL = replace(totalBRL,".",",")
		
	valorBRLGravidezUSD = 0
	valorBRLGravidezBRL = 0
	
	tipoGravidez = 0
	adGravidezN = "adGravidez"&pax_count
	if request.Form(adGravidezN) = 1 then
		valorBRLGravidezN = "valorBRLGravidez"&pax_count
		valorBRLGravidezUSD = 26.50
		valorBRLGravidezBRL = request.Form(valorBRLGravidezN)
		
		valorBRLGravidezUSD = replace(valorBRLGravidezUSD,".",",")
		valorBRLGravidezBRL = replace(valorBRLGravidezBRL,".",",")
		tipoGravidez = 1
	end if
	
	totalUSD = cdbl(totalUSD) + cdbl(valorBRLGravidezUSD)
	totalBRL = cdbl(totalBRL) + cdbl(valorBRLGravidezBRL)
	
	beneficiario_nomeN =  "beneficiario_nome"&pax_count
	beneficiario_nome = replace(Request.Form(beneficiario_nomeN),"'"," ")
	
	beneficiario_cpfN =  "beneficiario_cpf"&pax_count
	beneficiario_cpf = replace(Request.Form(beneficiario_cpfN),"'"," ")

	sobrePaxN =  "sobrePax"&pax_count
	sobrePax = replace(Request.Form(sobrePaxN),"'"," ")

	nomePaxN = "nomePax"&pax_count
	nomePax =  replace(Request.Form(nomePaxN),"'"," ")

	sexoPaxN = "sexoPax"&pax_count
	sexoPax = Request.Form(sexoPaxN)

	docPaxN = "docPax"&pax_count
	docPax = Request.Form(docPaxN)
 		
	emailPaxN = "emailPax"&pax_count
	emailPax = request.Form(emailPaxN)
	
	semanas_GestacaoN = "semanasGes"&pax_count
	semanas_Gestacao = request.Form(semanas_GestacaoN)

	DDDcelularPaxN = "DDDcelularPax"&pax_count
	DDDcelularPax = request.Form(DDDcelularPaxN)
	
	celularPaxN = "celularPax"&pax_count
	celularPax = request.Form(celularPaxN)
	
	tipoPaxN = "tipoPax"&pax_count
	tipoPax = Request.Form(tipoPaxN)

	dayN = "day"&pax_count
	dia = Request.Form(dayN)
	
	monthN = "month"&pax_count
	mes = Request.Form(monthN)

	yearN = "year"&pax_count
	ano = Request.Form(yearN)
 
	dtNascimento = dia &"/"& mes &"/"& ano	

	idadePaxN = "idadePax"&pax_count
	idadePax = Request.Form(idadePaxN)
	
	origemPassN = "origemPass"&pax_count
	origemPass = Request.Form(origemPassN)
	
	
	if request.Form("AcaoSubmit") <> 2 and request.Form("AcaoSubmit") <> 3 then 'impede erro caso o formulário nao seja totalmente cadastrado (Salvar e Sair)
		'35 < 37 E PLANO = 4901 
		if cint(regraTarifaRS("idadeMinima")) < cint(idadePax) and request.FORM("idPlano") = "4901" then 
			response.Write "<script>alert('Idade maior que permitida.'); history.go(-1)</script>"
		end if

		'50 < 51 E PLANO = 4898 
		if cint(regraTarifaRS("idadeMinima")) < cint(idadePax) and request.FORM("idPlano") = "4898" then 
			response.Write "<script>alert('Idade maior que permitida.'); history.go(-1)</script>"
		end if
	end if
		
	fileCliN = "fileCli"&pax_count
	fileCli = Request.Form(fileCliN)

	if acordo <> "1" then		 
		'calcular comissao desta emissao
		totalUSD = replace(totalUSD,".",",")
		totalBRL = replace(totalBRL,".",",")

		if rsCom.eof then
			response.Write("O sistema retornou um erro interno. <br> Detalhes do erro: <br> A Comissão do plano não está cadastrada, favor entrar em contato com Affinity")
			response.End()
		end if
		
		if not isnull(rsCom("porcentagem")) and rsCom("porcentagem")<>"" then comissao = rsCom("porcentagem") else comissao = 0
		
		if not isnull(rsCom("bola")) and rsCom("porcentagem")<>"" then bolaP = rsCom("bola") else bolaP = 0
		
		cambio = replace(cambio,".",",")
		cambio = formatNumber(cambio)
		if userRS.eof OR ISNULL(userRS("comissao")) then
			comissaoBRLPRO = 0 
		else
			comissaoBRLPRO = totalBRL * formatNumber(userRS("comissao")/100)
		end if
				
		comissaoUSD = totalUSD * (comissao/100)
		overUSD = totalUSD * (bolaP/100)
		netUSD = totalUSD - comissaoUSD
		
		if request.Form("AcaoSubmit") <> 2 and request.Form("AcaoSubmit") <> 3 then 'impede erro caso o formulário nao seja totalmente cadastrado (Salvar e Sair)
			if lcase(regraTarifaRS(1))="n" then
				if totalUSD <> 0 then totalBRL = totalUSD else totalBRL = totalBRL 'se é plano nacional, então força o BRL ser em reais
				comissaoBRL =  totalBRL * (comissao/100)
				overBRL = totalBRL * (bolaP/100)
				netBRL =  totalBRL - comissaoBRL
			else
				netBRL = netUSD * cambio
				overBRL = overUSD * cambio
				comissaoBRL = comissaoUSD * cambio 
			end if
		else
			netBRL = netUSD * cambio
			overBRL = overUSD * cambio
			comissaoBRL = comissaoUSD * cambio
		end if
		
		totalUSD = forMoeda(totalUSD,"2")
		totalBRL = forMoeda(totalBRL,"2")	

	else
		if destino = 4 then
			totalBRL = totalUSD
		end if
		netBRL = totalBRL
		valorliquidoBRL = netBRL
	end if		
					
	'desconto familiar, ZERA se nao for o primeiro
	if familiar = "1" AND familiarSeg <> 1 and valorBRLGravidezUSD = 0 and pax_count <> 1 then
		totalUSD = cdbl(valorBRLGravidezUSD)
		totalBRL = cdbl(valorBRLGravidezBRL)
		cambio = replace(cambio,",",".")
		comissaoUSD = 0
		overUSD = 0
		netUSD = 0
		comissaoBRL = 0
		overBRL = 0
		netBRL = 0
		valorLiquidoBRL = 0
		comissaoBRLPRO =  0
	end if
	
	'se houver cupom calcular desconto e abatimento de comissao
	if cupomDescontoId <> "0" and cupomDescontoId <> "" and isNumeric(cupomDescontoId) then

		desconto = getDesconto(cupomDescontoId, processo, idAge, planoId)

		if isNumeric(desconto) and desconto > 0 and desconto < 100 then

			comissao = replace(replace(comissao,".",""),",",".")
			bolaP = replace(replace(bolaP,".",""),",",".")

			totalUSDAnterior = replace(totalUSD, ".", ",")
			totalBRLAnterior = replace(totalBRL, ".", ",")

			'calculo da comissao deve ser em cima do valor antigo
			comissaoBRL = totalBRLAnterior * ((comissao - desconto) / 100)
			comissaoUSD = totalUSDAnterior * ((comissao - desconto) / 100)

			totalUSD = calculaDesconto(totalUSD, desconto)
			totalBRL = calculaDesconto(totalBRL, desconto)

			netUSD = totalUSD - comissaoUSD
			netBRL = totalBRL - comissaoBRL 

			overUSD = totalUSD * (bolaP/100)
			overBRL = totalBRL * (bolaP/100)
			

		end if

	end if
			
	'garantindo que zere os campos em dolar dos planos nacionais
	if lcase(regraTarifaRS(1)) = "n" then 
		totalUSD = 0
		netUSD = 0
		comissaoUSD = 0
	end if

	'se não for reserva carregada então dê INSERT, senão UPDATE
	if request.Form("carregarReserva") <> 1 then
		'hitorico do processo
		objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Insere PAX "&pax_count&" | Emissor "&emissor&"-"&emissorLogin&"')")

		if docPax = "" OR LEN(docPax) < 4 then
			response.write "Atencao: Documento obrigatorio nao informado (passageiro "&pax_count&") COD: 5049. Clique em voltar e informe o documento obrigatorio."
			response.End()
		end if

		' Trava para evitar voucher repetidos
		set verificaVoucherTemp = objConn.Execute("SELECT * FROM voucherTemp WHERE processoId='"&processo&"' AND seqPro='"&pax_count&"'")
						
		IF verificaVoucherTemp.eof then
			if idadePax = "NaN" then idadePax = ""										
				sql = "INSERT INTO voucherTemp (processoId, seqPro, dataEmissao, horaEmissao, cambio, meio, agencia, voucherAge, representante, plano, inicioVigencia,fimVigencia,dias,totalBRL,totalUSD,comissaoUSD,comissaoBRL,netUSD,netBRL,overUSD, overBRL, destino, nome, sobrenome, beneficiario_nome, beneficiario_cpf,  documento, email, celular, tipoDoc, endereco, fone, cidade, uf, cep, numero, bairro, complemento, idade, sexo, pais, familiar, emitido, emissor, emissorLogin, pagamento, inicioViagem,fimViagem,fileCli, acordo, planoAcordo, cancelado, clienteid, valorliquidoBRL, promotor1, promotor2, promotor3,comissaoBRL1, comissaoBRL2, dtNascimento, contatoNome, contatoFone,contatoEndereco, hotel, endHotel, foneHotel, planoId, numeroCartao, ccTitular, bolTipo, bolCPF, bolNome, bolEnd, bolBairro, bolCep, bolCidade, bolUF, bolEmail, tipoGravidez, semanas_gestacao) VALUES ("&processo&",'"&pax_count&"','"&data(dataEmissao,2,0)&"','"&horaEmissao&"','"&formoeda(cambio,2)&"','"&meio&"','"&cliRS("ni")&"','"&voucherAge&"','"&representante&"','"&plano&"','"&data(inicioViagem,2,0)&"','"&data(fimViagem,2,0)&"','"&dias&"','"&formoeda(totalBRL,2)&"','"&formoeda(totalUSD,2)&"','"&formoeda(comissaoUSD,2)&"','"&formoeda(comissaoBRL,2)&"','"&formoeda(netUSD,2)&"','"&formoeda(netBRL,2)&"','"&formoeda(overUSD,2)&"','"&formoeda(overBRL,2)&"','"&destino&"','"&replace(nomePax,"'","´")&"','"&replace(sobrePax,"'","´")&"','"&replace(beneficiario_nome,"'","´")&"','"&replace(beneficiario_cpf,"'","´")&"','"&docPax&"','"&emailPax&"','("&DDDcelularPax&")"&celularPax&"','"&tipoPax&"','"&enderecoPax&"','"&fone&"','"&cidadePax&"','"&ufPax&"','"&cepPax&"','"&numeroPax&"','"&bairroPax&"','"&complementoPax&"','"&idadePax&"','"&sexoPax&"','"&pais&"','"&familiar&"','1','"&emissor&"','"&emissorLogin&"','"&pagamento&"','"&inicioViagem2&"','"&fimViagem2&"','"&fileCli&"','"&acordo&"','"&planoAcordo&"',0,'"&Request.cookies("wlabel")("revId")&"','"&formoeda(valorliquidoBRL,2)&"','"&cliRS("promotor1")&"','"&cliRS("promotor2")&"','"&cliRS("promotor3")&"','"&formoeda(comissaoBRLPRO,2)&"','"&formoeda(comissaoBRLPRO,2)&"','"&dtNascimento&"','"&contatoNome&"','"&contatoFone&"','"&contatoEndereco&"','"&hotel&"','"&endHotel&"','"&foneHotel&"','"&planoId&"','"&request.form("numeroCartao")&"','"&request.Form("titular")&"','"&bolTipo&"','"&bolCPF&"','"&bolNome&"','"&bolEnd&"','"&bolBairro&"','"&bolCep&"','"&bolCidade&"','"&bolUF&"','"&bolEmail&"','"&tipoGravidez&"','"&semanas_Gestacao&"')"					

			
			objConn.Execute(sql)											
		END IF
	else
		if idadePax = "NaN" then idadePax = ""

		sql =       "UPDATE voucherTemp SET  "
		sql = sql & "processoId="&processo&", voucher='"&voucher&"', dataEmissao='"&data(dataEmissao,2,0)&"', horaEmissao='"&horaEmissao&"', cambio='"&formoeda(cambio,2)&"', "
		sql = sql & "meio='"&meio&"', agencia='"&cliRS("ni")&"', voucherAge = '"&voucherAge&"', representante = '"&representante&"', plano='"&plano&"',  "
		sql = sql & "inicioVigencia='"&data(inicioViagem,2,0)&"', fimVigencia='"&data(fimViagem,2,0)&"', dias='"&dias&"', totalBRL='"&formoeda(totalBRL,2)&"', totalUSD='"&formoeda(totalUSD,2)&"', "
		sql = sql & "comissaoUSD='"&formoeda(comissaoUSD,2)&"', comissaoBRL='"&formoeda(comissaoBRL,2)&"', netUSD='"&formoeda(netUSD,2)&"', netBRL='"&formoeda(netBRL,2)&"', overUSD='"&formoeda(overUSD,2)&"', overBRL='"&formoeda(overBRL,2)&"', "
		sql = sql & "destino='"&destino&"', nome='"&replace(nomePax,"'","´")&"', sobrenome='"&replace(sobrePax,"'","´")&"', beneficiario_nome='"&replace(beneficiario_nome,"'","´")&"', beneficiario_cpf='"&replace(beneficiario_cpf,"'","´")&"', documento='"&docPax&"', email='"&emailPax&"', "
		sql = sql & "celular='("&DDDcelularPax&")"&celularPax&"', tipoDoc='"&tipoPax&"', endereco='"&enderecoPax&"', fone='"&fone&"', cidade='"&cidadePax&"', uf='"&ufPax&"', "
		sql = sql & "cep='"&cepPax&"', numero='"&numeroPax&"', bairro='"&bairroPax&"', idade='"&idadePax&"', sexo='"&sexoPax&"', pais='"&pais&"', familiar='"&familiar&"', emitido='1', emissor='"&emissor&"', "
		sql = sql & "emissorLogin='"&emissorLogin&"', pagamento='"&pagamento&"', inicioViagem='"&inicioViagem2&"', fimViagem='"&fimViagem2&"', fileCli='"&fileCli&"', "
		sql = sql & "acordo='"&acordo&"', planoAcordo='"&planoAcordo&"', cancelado='0', clienteid='"&Request.cookies("FCNET_MIC")("idAge")&"', "
		sql = sql & "valorliquidoBRL='"&formoeda(valorliquidoBRL,2)&"', promotor1='"&cliRS("promotor1")&"', promotor2='"&cliRS("promotor2")&"', promotor3='"&cliRS("promotor3")&"', comissaoBRL1='"&formoeda(comissaoBRLPRO,2)&"', "
		sql = sql & "comissaoBRL2='"&formoeda(comissaoBRLPRO,2)&"', dtNascimento='"&dtNascimento&"', contatoNome='"&contatoNome&"', contatoFone='"&contatoFone&"', "
		sql = sql & "contatoEndereco='"&contatoEndereco&"', hotel='"&hotel&"', endHotel='"&endHotel&"', foneHotel='"&foneHotel&"', planoId='"&planoId&"', "
		sql = sql & "numeroCartao='"&request.form("numeroCartao")&"', ccTitular='"&request.form("titular")&"', "
		sql = sql & "bolTipo = '"&bolTipo&"', bolCPF='"&bolCPF&"', bolNome='"&bolNome&"', bolEnd='"&bolEnd&"', bolBairro='"&bolBairro&"', "
		sql = sql & "bolUF='"&bolUF&"', bolEmail='"&bolEmail&"', bolCep='"&bolCep&"', bolCidade='"&bolCidade&"', tipoGravidez = '"&tipoGravidez&"' "
		sql = sql & "WHERE processoId='"&processo&"' AND seqPro='"&pax_count&"'"
		objConn.Execute(sql)
	end if
		
	'hitorico do processo
	objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Atualiza PAX "&pax_count&"')")
		
	'sequencia do familiar'
	familiarSeg = familiarSeg + 1
	
	upgrade_cancel = Request.Form("planoCancelamento_id"&pax_count)
	upgrade_covid = Request.Form("planoCovid_id"&pax_count)

	if upgrade_cancel <> "" and upgrade_cancel <> 0 then
		Set rsUpgrade = objConn.Execute("select planos_upgrade_tipo.nome as nome_upgrade, preco, nacionalidade FROM planos_upgrade_tipo LEFT JOIN planos upgrade on upgrade.up_tipo_id = planos_upgrade_tipo.id LEFT JOIN valoresdiarios on valoresdiarios.planoId = upgrade.id where upgrade.id in (SELECT upGradeId from planos_upgrade where planoId = "&idPlano&") and upgrade.id = "&upgrade_cancel&" order by planos_upgrade_tipo.nome")

		if not rsUpgrade.eof then
			objConn.execute("INSERT INTO voucherTemp_apoio_upgrade (processoId,seqPro,upGradeId,valorUSD,valorBRL,tipo) VALUES ('"&processo&"','"&pax_count&"','"&upgrade_cancel&"','"&forMoeda(rsUpgrade("preco"),2)&"','"&forMoeda(rsUpgrade("preco")*rateBRL,2)&"','"&rsUpgrade("nome_upgrade")&"')")
		end if
	end if

	if upgrade_covid <> "" and upgrade_covid <> 0 then
		Set rsUpgrade = objConn.Execute("select planos_upgrade_tipo.nome as nome_upgrade, preco, nacionalidade FROM planos_upgrade_tipo LEFT JOIN planos upgrade on upgrade.up_tipo_id = planos_upgrade_tipo.id LEFT JOIN valoresdiarios on valoresdiarios.planoId = upgrade.id where upgrade.id in (SELECT upGradeId from planos_upgrade where planoId = "&idPlano&") and upgrade.id = "&upgrade_covid&" order by planos_upgrade_tipo.nome")

		if not rsUpgrade.eof then			
			Call CalcCovid (upgrade_covid, idPlano, destino, idadePax, dias)		

			if rsUpgrade("nacionalidade") = "n" then
				priceCovidBR = 0
			else
				priceCovidBR = priceCovid * rateBRL
			end if

			if rsUpgrade("nacionalidade") = "n" then
				objConn.execute("INSERT INTO voucherTemp_apoio_upgrade (processoId,seqPro,upGradeId,valorUSD,valorBRL,tipo) VALUES ('"&processo&"','"&pax_count&"','"&upgrade_covid&"','"&forMoeda(priceCovidBR,2)&"','"&forMoeda(priceCovid,2)&"','"&rsUpgrade("nome_upgrade")&"')")				
			else
				objConn.execute("INSERT INTO voucherTemp_apoio_upgrade (processoId,seqPro,upGradeId,valorUSD,valorBRL,tipo) VALUES ('"&processo&"','"&pax_count&"','"&upgrade_covid&"','"&forMoeda(priceCovid,2)&"','"&forMoeda(priceCovidBR,2)&"','"&rsUpgrade("nome_upgrade")&"')")
			end if
		end if
	end if
	
	' verificar se tem upgrade 		
	set upgradePaxRS = objConn.execute("SELECT * from voucherTemp_apoio_upgrade where processoId= '"&processo&"' and seqPro = '"&pax_count&"' ")

	WHILE not upgradePaxRS.EOF						
		Set rsCom = objConn.Execute("SELECT * FROM comissao where idAge ='"&idAge&"' AND idPlano='"&upgradePaxRS("upGradeId")&"'")
		
		if not rsCom.eof then
			if not isnull(rsCom("porcentagem")) and rsCom("porcentagem")<>"" then comissao = rsCom("porcentagem") else comissao = 0
			if not isnull(rsCom("bola")) and rsCom("porcentagem")<>"" then bolaP = rsCom("bola") else bolaP = 0
		else
			comissao = 0
			bolaP = 0
		end if						
		
		voucherAge = voucherAge + 1
		comissaoUSD = upgradePaxRS("valorUSD") * (comissao/100)
		comissaoBRL = upgradePaxRS("valorBRL") * (comissao/100)
		netUSD = upgradePaxRS("valorUSD") - comissaoUSD
		netBRL = upgradePaxRS("valorBRL") - comissaoBRL
		overUSD = upgradePaxRS("valorUSD") * (bolaP/100)
		overBRL = upgradePaxRS("valorBRL") * (bolaP/100)
		upTotalUSD = upgradePaxRS("valorUSD")
		upTotalBRL = upgradePaxRS("valorBRL")

		'regra cupom no upgrade
		if cupomDescontoId <> "0" and cupomDescontoId <> "" and isNumeric(cupomDescontoId) then

			desconto = getDesconto(cupomDescontoId, processo, idAge, planoId)

			if isNumeric(desconto) and desconto > 0 and desconto < 100 then				

				'calculo da comissao deve ser em cima do valor antigo
				comissaoBRL = upgradePaxRS("valorBRL") * ((comissao - desconto) / 100)
				comissaoUSD = upgradePaxRS("valorUSD") * ((comissao - desconto) / 100)

				upTotalUSD = calculaDesconto(upgradePaxRS("valorUSD"), desconto)
				upTotalBRL = calculaDesconto(upgradePaxRS("valorBRL"), desconto)

				netUSD = upTotalUSD - comissaoUSD
				netBRL = upTotalBRL - comissaoBRL 

				overUSD = uptotalUSD * (bolaP/100)
				overBRL = uptotalBRL * (bolaP/100)
			end if

		end if
		set testeUPRS = objconn.execute("SELECT * FROM voucherTemp where processoId= '"&processo&"' and seqPro = '"&pax_count&"' and planoId="&upgradePaxRS("upGradeId"))
		if testeUPRS.EOF then
		
			set nPlanoRS = objConn.execute("SELECT nPlano from planos where id="&upgradePaxRS("upGradeId"))
			sql = "INSERT INTO voucherTemp (processoId, seqPro, flagUp, dataEmissao, horaEmissao, cambio, meio, agencia, voucherAge, representante, plano, inicioVigencia,fimVigencia,dias,totalBRL,totalUSD,comissaoUSD,comissaoBRL,netUSD,netBRL,overUSD, overBRL, destino, nome, sobrenome, beneficiario_nome, beneficiario_cpf, documento, email, celular, tipoDoc, endereco, fone, cidade, uf, cep, idade, sexo, pais, familiar, emitido, emissor, emissorLogin, pagamento, inicioViagem,fimViagem,fileCli, acordo, planoAcordo, cancelado, clienteid, valorliquidoBRL, promotor1, promotor2, comissaoBRL1, comissaoBRL2, dtNascimento, contatoNome, contatoFone,contatoEndereco, planoId, semanas_gestacao, numero, bairro, complemento) VALUES ("&processo&",'"&pax_count&"','1','"&data(dataEmissao,2,0)&"','"&horaEmissao&"','"&formoeda(cambio,2)&"','"&meio&"','"&cliRS("ni")&"','"&voucherAge&"','"&representante&"','"&nPlanoRS(0)&"','"&data(inicioViagem,2,0)&"','"&data(fimViagem,2,0)&"','"&dias&"','"&formoeda(upTotalBRL,2)&"','"&formoeda(upTotalUSD,2)&"','"&formoeda(comissaoUSD,2)&"','"&formoeda(comissaoBRL,2)&"','"&formoeda(netUSD,2)&"','"&formoeda(netBRL,2)&"','"&formoeda(overUSD,2)&"','"&formoeda(overBRL,2)&"','"&destino&"','"&replace(nomePax,"'","´")&"','"&replace(sobrePax,"'","´")&"','"&replace(beneficiario_nome,"'","´")&"','"&replace(beneficiario_cpf,"'","´")&"','"&docPax&"','"&emailPax&"','("&DDDcelularPax&")"&celularPax&"','"&tipoPax&"','"&enderecoPax&"','"&fone&"','"&cidadePax&"','"&ufPax&"','"&cepPax&"','"&idadePax&"','"&sexoPax&"','"&pais&"','0','1','"&emissor&"','"&emissorLogin&"','"&pagamento&"','"&inicioViagem2&"','"&fimViagem2&"','"&fileCli&"','"&acordo&"','"&planoAcordo&"',0,'"&Request.cookies("wlabel")("revId")&"','"&formoeda(valorliquidoBRL,2)&"','"&cliRS("promotor1")&"','"&cliRS("promotor2")&"','"&formoeda(comissaoBRLPRO,2)&"','"&formoeda(comissaoBRLPRO,2)&"','"&dtNascimento&"','"&contatoNome&"','"&contatoFone&"','"&contatoEndereco&"','"&upgradePaxRS("upGradeId")&"','"&semanas_Gestacao&"','"&numeroPax&"','"&bairroPax&"','"&complementoPax&"' )"
			objConn.Execute(sql)
		end if
	
		upgradePaxRS.MOVENEXT
	WEND

NEXT

if idAge = "28" then
	
	'objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Teste envio de email')")
	'response.Redirect("emitir1.asp?AcaoSubmit="&request.Form("AcaoSubmit")&"&tavola="&tavola&"&processo="&request.Form("processo")&"&pagamento="&pagamento&"&ccMarca=""&tid="&request.Form("tid")&"&autorizacao="&request.Form("autorizacao")&"&emailTitular="&request.Form("emailTitular")&"&pg="&request.Form("pg")&"&AcaoSubmitBoleto="&request.Form("AcaoSubmitBoleto"))

end if 

'atualizar o valor total BRL do processo
set valorProcessoRS = objConn.Execute("SELECT SUM(totalBRL) from voucherTemp where processoId='"&processo&"'")
objConn.Execute("UPDATE emissaoProcesso set valorTotalBRL = '"&forMoeda(valorProcessoRS(0),2)&"',pgtoAprovado='4' WHERE id ='"&processo&"'")
valorProcessoRS.close
set valorProcessoRS = nothing 

	'hitorico do processo
	objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Direcionado para processamento de cartao')")
	'Fecha o Objeto de Conexão
	objConn.close
	Set objConn = Nothing 

	response.Redirect "pagamento.asp?processo="&request.Form("processo")&"&emailTitular="&request.Form("emailTitular")             
%>