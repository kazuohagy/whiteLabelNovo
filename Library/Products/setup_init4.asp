<!--#include file="../Common/micMainCon.asp" -->
<!--#include file="../Common/funcoes.asp" -->
<!--#include file="PriceFunctions.asp" -->

<%
	Session.CodePage=65001

plano_id = request.QueryString("planoId")
categoria_id = request.QueryString("categoria")
destino_id = request.QueryString("destino")
data_inicio	= cdate(request.QueryString("inicioViagem"))
data_fim = cdate(request.QueryString("fimViagem"))
idadeMenor = Request.QueryString("idadeMenor")
idadeMaior = Request.QueryString("idadeMaior")
cancelamento_id = request.QueryString("planoCancelamento")
covid_id = request.QueryString("planoCovid")
familiar = request.QueryString("familiar")
nPax = cint(idadeMenor) + cint(idadeMaior)
reg_usuario = request.cookies("FCNET_MIC")("login")
clienteId = request.cookies("FCNET_MIC")("idAge")
ip = request.ServerVariables("REMOTE_ADDR")
fone = request.QueryString("fone")
email = request.QueryString("email")
site = request.ServerVariables("HTTP_HOST")
address = request.QueryString("address")

if categoria_id = "20" then
	idadeMenor = 0
end if


if categoria_id = "" or destino_id = "" or data_inicio = "" or data_fim = "" or idadeMenor = "" or idadeMaior = "" then
	response.write "<script>alert('\Preencha todos os campos!')</script>"
	response.write "<script>window.history.back()</script>"
	objConn.close
	response.End()
end if

dias = DateDiff("d",data_inicio,data_fim) + 1

if idadeMenor = 0 and idadeMaior = 0 then
	response.write "<script>alert('\É necessário pelo menos 1 passageiro')</script>"
	response.write "<script>window.history.back()</script>"
	objConn.close
	response.End()
end if

if familiar = "" then familiar = 0 else familiar = "1"

if idadeMaior <> 0 and familiar = 1 then
	response.write "<script>alert('\Não é permitido pessoas com mais de 65 anos no plano friends')</script>"
	response.write "<script>window.history.back()</script>"
	objConn.close
	response.End()
end if

Dim vet_idadePax()
Redim vet_idadePax(npax-1)
counter = 0

for i = 0 to idadeMenor - 1	
	vet_idadePax(counter) = 20
	counter = counter + 1
next

for i = 0 to idadeMaior - 1
	vet_idadePax(counter) = 70
	counter = counter + 1
next

vSQLUpd = ""
vSQLUpd = vSQLUpd + "   reg_usuario 		= '"&reg_usuario&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   clienteId 			= '"&clienteId&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   ip 					= '"&ip&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   fone 				= '"&fone&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   email 				= '"&email&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   categoria_id		= '"&categoria_id&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   destino_id			= '"&destino_id&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   data_inicio			= '"&data(data_inicio,2,0)&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   data_fim			= '"&data(data_fim,2,0)&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   vigencia			= '"&dias&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   familiar_flag       = '"&familiar&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   nPax_total			= '"&nPax&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   nPax_Idoso			= '"&idadeMaior&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   nPax_Novo			= '"&idadeMenor&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   site        		= '"&site&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   upgradeCancel       = '"&cancelamento_id&"' ;"+chr(0)
vSQLUpd = vSQLUpd + "   upgradeCovid       = '"&covid_id&"' ;"+chr(0)

strSQL = "   INSERT INTO cotacao_reg "
strSQL = strSQL + "      " + InsertUpdate(vSQLUpd, 0) + " "
objConn.execute(strSQL)

set cotacaoRS = objConn.execute("SELECT TOP 1 id FROM cotacao_reg WHERE ip = '"&ip&"' order by id desc")
	cotacao_id	= cotacaoRS("id")
set cotacaoRS = NOTHING

if familiar = "1" then sql_add_fam = " AND familiar = '1' and minPaxFam <= '"&nPax&"' and maxPaxFam >= '"&nPax&"' " else sql_add_fam = " AND coalesce(familiar,0) = '0' "

if idadeMaior > "0" then sql_idoso = "AND idadeMinima > 65" else sql_idoso = " "

if plano_id = "" then	
	SQL = 	" SELECT id, nome, nacionalidade, CASE WHEN nacionalidade = 'i' THEN 'US$' ELSE 'R$' END AS MOEDA FROM planos WHERE coalesce(ageid,0)=0 AND coalesce(publicado,1)=1 "&sql_add_fam&" "&sql_idoso&" AND id IN (SELECT planoId FROM categoriaXPlano WHERE categoriaId = '"&categoria_id&"') AND id IN (SELECT planoId FROM viagem_destinoPlano WHERE destinoId = '"&destino_id&"') AND id IN (SELECT planoId FROM valoresdiarios WHERE dias = '"&dias&"') AND versaotarifa = '4' AND vigenciaMaxima >= '"&dias&"' ORDER BY id "

	response.write SQL
	

	set cotacaoRS = objConn.execute(SQL)

	if cotacaoRS.EOF then
		response.write "<script>alert('\Não foi encontrado nenhum plano com esse perfil!')</script>"
		response.write "<script>window.history.back()</script>"	
		set cotacaoRS = nothing
		objConn.CLOSE
		response.end
	end if
else	
	set cotacaoRS =	objConn.execute("SELECT * ,  CASE WHEN nacionalidade = 'i' THEN 'US$' ELSE 'R$' END AS MOEDA FROM planos WHERE id IN (SELECT planoId FROM valoresdiarios WHERE dias <= '"&dias&"' AND planoId IN ("& plano_id &")) AND nPlano IN (SELECT plano FROM viagem_destinoPlano WHERE plano IN (SELECT nPlano FROM planos WHERE id IN ("&plano_id&")) and destinoId = '"&destino_id&"') "&sql_add_fam&" "&sql_idoso&" AND vigenciaMaxima >= '"&dias&"' ORDER BY ordemExibicao")		
	
	if cotacaoRS.EOF then
		response.write "<script>alert('\Passe planos válidos!')</script>"
		response.write "<script>window.history.back()</script>"	
		set cotacaoRS = nothing
		objConn.CLOSE
		response.end
	end if	

end if

response.write SQL

response.end()
While not cotacaoRS.EOF	
	SQL = "INSERT INTO cotacao_reg_pax (cotacao_id , plano_id , plano_nome, sequencia_pax, tarifa_USD, tarifa_BRL, tarifa_original, fator_venda, moeda, idade, idade_pax, familiar, tarifa_upgradeCancel, tarifa_upgradeCovid, tarifa_upgradeCovidOriginal) "

	Call CalcPrice(cotacaoRS("id"), dias, destino_id, vet_idadePax, familiar,categoria_id)		
	
	if cancelamento_id <> "" then
		SQLCANCEL = "SELECT preco FROM planos_upgrade inner join valoresdiarios on upGradeId = valoresdiarios.planoId WHERE planos_upgrade.planoId = '"& cotacaoRS("id") &"' and upGradeId = '"& cancelamento_id &"'"
		set cancelRS = objConn.execute(SQLCANCEL)					
		
		if not cancelRS.eof then		
			cancelPrice = forMoeda(cancelRS("preco"),2)
		end if
	end if	

	SQLValue = "VALUES "
	For i = 0 To nPax - 1	  
		if covid_id <> "" then
			Call CalcCovid(covid_id, cotacaoRS("id"), destino_id, vet_idadePax(i), dias)
		end if

		SQLValue = SQLValue & " ('"& cotacao_id &"','"& cotacaoRS("id") &"','"& cotacaoRS("nome") &"','"& i+1 &"','"& forMoeda(pricePax(i),2) &"','"& forMoeda(pricePaxBR(i),2) &"','"& forMoeda(priceOriginal(i),2) &"', 1 ,'"& cotacaoRS("moeda") &"','"& CINT(vet_idadePax(i)) &"','"& CINT(vet_idadePax(i)) &"','"& familiar &"','"& cancelPrice &"','"& forMoeda(priceCovid,2) &"','"& forMoeda(priceCovidOriginal,2) &"'),"
	Next
	SQLValue = Left(SQLValue, Len(SQLValue)-1)	
	objConn.execute(SQL & SQLValue)	
cotacaoRS.MOVENEXT
WEND

set cotacaoRS = NOTHING
objConn.cLOSE

response.Redirect address & "cotacao_id=" & cotacao_id
%>

