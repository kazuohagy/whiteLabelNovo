<!--#include file="../Common/micMainCon.asp" -->
<!--#include file="../Common/funcoes.asp" -->
<!--#include file="../Common/cambio.asp" -->

<%  
Session.CodePage=65001

Dim parcela

function achaTarifa_PUB(nPlano,vigencia,idade)
	Dim planoRS, cotacaoCustoRS, tarifaPax, diferencaVigencia, erro

	erro = 0
	
	Set tar_planoRS = objConn.Execute("SELECT * FROM planos where nPlano='"&nPlano&"' and COALESCE(ageId,0) = '0' and COALESCE(publicado,0)='1'")
	if tar_planoRS.EOF then
		Set tar_planoRS = objConn.Execute("SELECT * FROM planos where nPlano='"&nPlano&"' and COALESCE(ageId,0) = '0' and COALESCE(operador,0)='1'")
	end if	

	if tar_planoRS.EOF then
		Set tar_planoRS = objConn.Execute("SELECT top 1 * FROM planos where nPlano='"&nPlano&"' order by id")
	end if
		
	if (CINT(tar_planoRS("limiteAdicional")) >= CINT(vigencia)) then 
		Set cotacaoCustoRS = objConn.Execute("SELECT * FROM valoresdiarios where planoId='"&tar_planoRS("id")&"' and dias='"&vigencia&"'")

		if cotacaoCustoRS.EOF  then
			achaTarifa_PUB = "0"
		else
			achaTarifa_PUB = cotacaoCustoRS("preco")
		end if		
	else				
		diferencaVigencia = (vigencia - tar_planoRS("limiteAdicional")) 
			
		Set cotacaoCustoRS = objConn.Execute("SELECT * FROM valoresdiarios where planoId='"&tar_planoRS("id")&"' and dias='"&tar_planoRS("limiteAdicional")&"'")
		
		achaTarifa_PUB = (tar_planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco")
	end if	
end function  


function trataNulo(valor)
	if isnull(valor) then
		valor = ""
	else
		valor = TRIM(valor)
	end if
	trataNulo = valor
end function

voucher = request.QueryString("voucher")
id		= request.QueryString("id")
idioma	= request.QueryString("idioma")

if idioma = "" then idioma = "1"

if request.QueryString("voucher") <> "" then
	set voucherRS = objConn.execute("SELECT * FROM voucher where voucher='"&request.QueryString("voucher")&"'")' and id='"&request.QueryString("id")&"' and coalesce(cancelado,0) = 0")
elseif request.QueryString("id") <> "" then
	set voucherRS = objConn.execute("SELECT * FROM voucher where id='"&request.QueryString("id")&"'")' and id='"&request.QueryString("id")&"' and coalesce(cancelado,0) = 0")
end if

if voucherRS.eof then
	response.write "Certificado nao localizado"
	response.End()
end if
    
if voucherRS("cancelado") = "1" then
	response.write "Certificado cancelado"
	response.End()
end if   

if voucherRS("flagUp") = "1" then
	set voucherARS = objConn.execute("SELECT voucher from voucher where coalesce(flagUp,0)=0 and processoId = '"&voucherRS("processoId")&"' and sobrenome='"&voucherRS("sobrenome")&"' and nome='"&voucherRS("nome")&"'")
	
	if not voucherARS.EOF then response.redirect "geraCertificadoCHUBB.asp?voucher="&voucherARS("voucher")&"&idioma="&idioma&"&anexo="&request.QueryString("anexo")

	set voucherARS = nothing
end if

documento = voucherRS("documento")

'' verifica se tem upgrade
sql_cob_upgrade = ""
set voucherARS = objConn.execute("SELECT plano from voucher where coalesce(flagUp,0)=1 and processoId = '"&voucherRS("processoId")&"' and sobrenome='"&voucherRS("sobrenome")&"' and nome='"&voucherRS("nome")&"'")
if not voucherARS.EOF then sql_cob_upgrade = " OR nPlano = '"&voucherARS(0)&"' "

set voucherARS = nothing''''
	
'' verifica se tem upgrade de covid
	sql_cob_upgradeCov = ""
	set voucherCovARS = objConn.execute("SELECT plano from voucher where plano like '7%' and coalesce(flagUp,0)=1 and processoId = '"&voucherRS("processoId")&"' and sobrenome='"&voucherRS("sobrenome")&"' and nome='"&voucherRS("nome")&"'")
	if not voucherCovARS.EOF then sql_cob_upgradeCov = " '"&voucherCovARS(0)&"' " else sql_cob_upgradeCov = "'0'" end if
	
	set voucherCovARS = nothing''''	
    
set planoRS			= objConn.execute("SELECT * FROM planos where id='"&voucherRS("planoId")&"' ")

if planoRS.EOF then
	response.write "Plano nao identificado"
	response.End()
end if

id		= voucherRS("id")

if voucherRS("acordo") <> "1" or voucherRS("familiar") = "1" then

if right(voucherRS("voucher"),2) = "F1" OR right(voucherRS("voucher"),2) = "F2" OR right(voucherRS("voucher"),2) = "F3" OR right(voucherRS("voucher"),2) = "F4"  OR right(voucherRS("voucher"),2) = "F5"  OR right(voucherRS("voucher"),2) = "F6" then
familiar = 1
	' familiar
	'set total_processoRS = objConn.execute("SELECT coalesce(sum(totalBRL),0) as valorTotalBRL, coalesce(sum(totalUSD),0) as valorTotalUSD  from VOUCHER where processoid = "&voucherRS("processoId"))
	set processoRS = objConn.execute("SELECT nPax, familiar,parcelas from emissaoprocesso where id = "&voucherRS("processoId"))
	if UCASE(planoRS("nacionalidade")) = "I" then
		premioBRL = achaTarifa_PUB(voucherRS("plano"),voucherRS("dias"),voucherRS("idade")) * voucherRS("cambio")
	else
		premioBRL = achaTarifa_PUB(voucherRS("plano"),voucherRS("dias"),voucherRS("idade"))
	end if
	premioUSD = achaTarifa_PUB(voucherRS("plano"),voucherRS("dias"),voucherRS("idade"))
	plano_N = voucherRS("plano")
	
		premioBRL = premioBRL/processoRS("nPax")
		premioUSD = premioUSD/processoRS("nPax")
		
		set tempRS = objConn.execute("SELECT top 1 plano FROM voucher WHERE processoId='"&voucherRS("processoId")&"' order by id")
		plano_N = tempRS(0)
else
	familiar = 0
	premioBRL = voucherRS("totalBRL")
	premioUSD = voucherRS("totalUSD")
	plano_N = voucherRS("plano")
	
end if

else ' para planos acordo
	if UCASE(planoRS("nacionalidade")) = "I" then
		premioBRL = achaTarifa_PUB(voucherRS("plano"),voucherRS("dias"),voucherRS("idade")) * voucherRS("cambio")
	else
		premioBRL = achaTarifa_PUB(voucherRS("plano"),voucherRS("dias"),voucherRS("idade"))
	end if
	premioUSD = achaTarifa_PUB(voucherRS("plano"),voucherRS("dias"),voucherRS("idade"))
	plano_N = voucherRS("plano")
'achaTarifa_PUB


end if
premioBRL_cob = premioBRL



''' verificar versao de cobertura

set versaoCOBRS = objConn.execute("SELECT versao FROM coberturasPlanos_versao where validade_de < '"&data(voucherRS("dataEmissao"),2,0)&"' AND validade_ate > '"&data(voucherRS("dataEmissao"),2,0)&"'")
versaoCob = versaoCOBRS(0)
'set versaoCOBRS = nothing




'chubb next 

	certificado_modelo = "bilhete_chubb_2021_nxt.pdf"
	LinkCG = "https://www.nextseguroviagem.com.br/next-cg/"
	'atualizar qrCode
	qrcodelink = Server.MapPath("qr-code/next/qr-next-2021.png")
	

 



if IsNumeric(voucherRS("destino")) then

	set rsdetinoCombo = objConn.execute("select * from viagem_destino  where id='"&voucherRS("destino")&"'")
	destino = ucase(rsdetinoCombo("nome"))
else
	destino = ucase(voucherRS("destino"))
end if



if idioma = "2" then certificado_modelo = REPLACE(certificado_modelo,".pdf","_EN.pdf")

if idioma = "3" then certificado_modelo = REPLACE(certificado_modelo,".pdf","_ES.pdf")


Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.OpenDocument( Server.MapPath("modelo/"&certificado_modelo) )
Set Font = Doc.Fonts("Arial")
' Obtain the only page's canvas
Set Page = Doc.Pages(1)
'resolve o problema das letras invertidas
Page.ResetCoordinates
Set Canvas = Page.Canvas




	set tempLogoRS = nothing

'''' fim logo agencia



'voucher
Set Param = Pdf.CreateParam("x=100; y=668; Size=9; color=black;")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("voucher"), Param, Font

'data emissao
Set Param = Pdf.CreateParam("x=475; y=668; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("regdata") , Param, Font

'pax
Set Param = Pdf.CreateParam("x=65; y=615; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText TRIM(voucherRS("nome"))& " " & TRIM(voucherRS("sobrenome")) , Param, Font

'documento
Set Param = Pdf.CreateParam("x=373; y=615; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText TRIM(voucherRS("documento")), Param, Font

'dtNascimento
Set Param = Pdf.CreateParam("x=525; y=615; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("dtNascimento"), Param, Font

'endereco
Set Param = Pdf.CreateParam("x=74; y=602; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("endereco") & " " & voucherRS("numero"), Param, Font

'bairro
Set Param = Pdf.CreateParam("x=341; y=602; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText trataNulo(voucherRS("bairro")), Param, Font

'cidade
Set Param = Pdf.CreateParam("x=65; y=590; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("cidade"), Param, Font

'uf
Set Param = Pdf.CreateParam("x=341; y=590; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("uf"), Param, Font

'cep
Set Param = Pdf.CreateParam("x=500; y=590; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText trataNulo(voucherRS("cep")), Param, Font

'plano
Set Param = Pdf.CreateParam("x=65; y=576; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText planoRS("nPlano") & "-" & planoRS("nome") , Param, Font

'destino
Set Param = Pdf.CreateParam("x=360; y=576; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText UCASE(destino) , Param, Font

'inicioVigencia
Set Param = Pdf.CreateParam("x=133; y=563; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText CDATE(voucherRS("inicioVigencia")), Param, Font

'fimVigencia
Set Param = Pdf.CreateParam("x=420; y=563; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText CDATE(voucherRS("fimVigencia")), Param, Font

'vigencia
Set Param = Pdf.CreateParam("x=530; y=563; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText voucherRS("dias"), Param, Font

set addRS = objConn.EXECUTE("SELECT COALESCE(SUM(totalUSD),0) as totalUSD, COALESCE(SUM(totalBRL),0) as totalBRL  FROM voucher where processoid='"&voucherRS("processoid")&"' and sobrenome='"&REPLACE(voucherRS("sobrenome"),"'","")&"' and voucher <> '"&voucherRS("voucher")&"' and flagUp = '1' and nome='"&voucherRS("nome")&"'")
if not addRS.EOF then
	premioBRL = premioBRL + addRS("totalBRL")
	if UCASE(planoRS("nacionalidade")) = "I" then
	premioUSD = premioUSD + addRS("totalUSD")
	end if
end if

'tarifa USD
Set Param = Pdf.CreateParam("x=133; y=550; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText formatNumber(premioUSD,2), Param, Font


'cambio
Set Param = Pdf.CreateParam("x=365; y=550; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText formatNumber(voucherRS("cambio"),2), Param, Font



'tarifa BRL
Set Param = Pdf.CreateParam("x=480; y=550; Size=9; color=black")
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText formatNumber(premioBRL,2), Param, Font

' forma pgto
Set Param = Pdf.CreateParam("x=133; y=537; Size=9; color=black")
if voucherRS("pagamento") = "CC" then
pagamento = "Cart. de Cred."
else
pagamento = "a vista"
end if
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText pagamento , Param, Font

' prazo pgto
Set Param = Pdf.CreateParam("x=395; y=537; Size=9; color=black")
	set prazoRS = objConn.execute("SELECT parcelas from emissaoprocesso where id = "&voucherRS("processoId"))
	parcelas = prazoRS("parcelas")
	set prazoRS = nothing
if parcelas = "" then parcelas = "1"
'corrige as posicoes do y nas apolices Next
 Param("y") = Param("y") + 39
Canvas.DrawText parcelas&"x" , Param, Font

if voucherRS("semanas_gestacao") <> "" then
Set Param = Pdf.CreateParam("x=27; y=555; Size=9; color=black")
Canvas.DrawText "Passageira declara "& voucherRS("semanas_gestacao") &" semanas de gestacao" , Param, Font
end if




if not isnull(planoRS("OBS_SEG_IDADE")) and 1 <> 1 then 

Set Param = Pdf.CreateParam("x=32; y=114; Size=9; color=gray; width=390; height=200")
Canvas.DrawText planoRS("OBS_SEG_IDADE") , Param, Font

end if
'''' fator de cobertura inexistente
SQL = "  DECLARE @total as float, @n_cob as float, @dif as float, @fator as float " & _

"		 select  'total' = sum(pesohdi), " & _
"		 'n_cob' = count(id), " & _
"		 'dif' = 100 - sum(pesohdi), " & _
"		 'fator'= (100 - sum(pesohdi))  / count(id) " & _
		
"		 FROM coberturas where tipo = '1' and id in (SELECT coberturaid from coberturasplanos " & _

"		 								where   pesohdi <> 0 and ( " & _
"											coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"'  )"&sql_cob_upgrade&" ) " & _
"									    	OR " & _
"											coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"' )"&sql_cob_upgrade&") " & _
"										)) "

'response.write SQL
'response.end


set fatorRS = objConn.execute(SQL)
''''''''''''''''''''''''''''''''''''
'response.write SQL & "<BR><BR>"
'response.End()
if idioma = "1" then
	descritivo_cob = "descritivo"
	simbolo = "simbolo"
	valor_cob = "valor"
else
	descritivo_cob = "eng_descritivo"
	simbolo = "simbolo_ingles"
	valor_cob = "valor_ingles"

end if

if idioma = "3" then
	descritivo_cob = "esp_descritivo"
	simbolo = "simbolo_espanhol"
	valor_cob = "valor_espanhol"
end if

'''' montar cobertura de assistencia
SQL = "  select " & _
"		 coberturas."&descritivo_cob&"  as descritivo, " & _
"		 coalesce(coberturas.pesoHdi,0) as pesoHdi, " & _
"		 coberturasplanos."&simbolo&"   as simbolo, " & _
"		 coberturasplanos."&valor_cob&" as valor " & _

"		FROM coberturas " & _

" 		inner join coberturasplanos ON coberturas.id = coberturasplanos.coberturaID " & _

"       where tipo='2' and coberturasplanos.versao_id = '"&versaoCob&"' and ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"' )  )" & _
"			or	coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 "&sql_cob_upgrade&"  ) )" & _
"				OR tipo='2' and coberturasplanos.versao_id = '"&versaoCob&"' and ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 0  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"' ) "&sql_cob_upgrade&" ) " & _
"			  )  order by coberturas.ordem "

set cobRS = objConn.execute(SQL)

Set param = Pdf.CreateParam("rows=10, cols=2, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 4
linha     = 0
conta_cob = 0

WHILE NOT cobRS.EOF and max_linha > linha

linha = linha + 1 
conta_cob = conta_cob + 1

if cobRS("pesoHDI") <> 0 then
	fator = fatorRS("fator")
else
	fator = 0
end if

	Table(linha,1).AddText TRIM(cobRS("descritivo")), "size=8; expand=true; color=black", Font
	Table(linha,1).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,1).Width   = "160"

	Table(linha,2).AddText TRIM(cobRS("simbolo"))&TRIM(cobRS("valor")), "size=8; expand=true; color=black", Font
	Table(linha,2).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,2).Width   = "114"
	
	'Table(linha,3).AddText "R$" & formatNumber(premioBRL*((cobRS("pesoHDI")+fator)/100),4), "size=8; expand=true; color=black", Font
	'Table(linha,3).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	'Table(linha,3).Width   = "48"

cobRS.MOVENEXT
WEND
eixoY = 470
 eixoY = eixoY + 39
Canvas.DrawTable Table, "x=27, y="&eixoY


Set param = Pdf.CreateParam("rows=10, cols=2, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 4
linha     = 0
'conta_cob = 0

WHILE NOT cobRS.EOF and max_linha > linha

linha = linha + 1 
conta_cob = conta_cob + 1

if cobRS("pesoHDI") <> 0 then
	fator = fatorRS("fator")
else
	fator = 0
end if

	Table(linha,1).AddText TRIM(cobRS("descritivo")), "size=8; expand=true; color=black", Font
	Table(linha,1).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,1).Width   = "160"

	Table(linha,2).AddText TRIM(cobRS("simbolo"))&TRIM(cobRS("valor")), "size=8; expand=true; color=black", Font
	Table(linha,2).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,2).Width   = "114"
	
	'Table(linha,3).AddText "R$" & formatNumber(premioBRL*((cobRS("pesoHDI")+fator)/100),4), "size=8; expand=true; color=black", Font
	'Table(linha,3).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	'Table(linha,3).Width   = "48"

cobRS.MOVENEXT
WEND

eixoY = eixoY + 2
Canvas.DrawTable Table, "x=300, y="&eixoY&""
'servicos de seguro 
SQL = "  select " & _
"		 coberturas."&descritivo_cob&"  as descritivo, " & _
"		 coalesce(coberturas.pesoHdi,0) as pesoHdi, " & _
"		 coberturasplanos."&simbolo&"   as simbolo, " & _
"		 coberturasplanos."&valor_cob&" as valor " & _

"		FROM coberturas " & _

" 		inner join coberturasplanos ON coberturas.id = coberturasplanos.coberturaID " & _

"       where ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"'  )    )  " & _

"			    OR " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and coalesce(nPlano,0) = "&sql_cob_upgradeCov&"   order by nplano  )  " & _
"			    OR " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 0  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"'  )  "&sql_cob_upgrade&" ) " & _
"			  ) and tipo='1'  and coberturasplanos.versao_id = '"&versaoCob&"'  order by coberturas.ordem "


set cobRS = objConn.execute(SQL)

Set param = Pdf.CreateParam("rows=17, cols=2, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 17
linha     = 0
conta_cob = 0

WHILE NOT cobRS.EOF and max_linha > linha

linha = linha + 1 
conta_cob = conta_cob + 1

if cobRS("pesoHDI") <> 0 then
	fator = fatorRS("fator")
else
	fator = 0
end if

	Table(linha,1).AddText TRIM(cobRS("descritivo")), "size=8; expand=true; color=black", Font
	Table(linha,1).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,1).Width   = "160"

	Table(linha,2).AddText TRIM(cobRS("simbolo"))&TRIM(cobRS("valor")), "size=8; expand=true; color=black", Font
	Table(linha,2).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,2).Width   = "114"
	
	'Table(linha,3).AddText "R$" & formatNumber(premioBRL*((cobRS("pesoHDI")+fator)/100),4), "size=8; expand=true; color=black", Font
	'Table(linha,3).AddText "", "size=8; expand=true; color=black", Font
	'Table(linha,3).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	'Table(linha,3).Width   = "48"

cobRS.MOVENEXT
WEND

'condicoes gerais

ARQ_CG = LinkCG



Set Param = Pdf.CreateParam("x=87; y=57; Size=8; color=orange")
 Param("x") = Param("x") + 90
Canvas.DrawText "Cond. Gerais: " & ARQ_CG , Param, Font

'' pagina 2
Set Page = Doc.Pages(2)
'resolve o problema das letras invertidas
Page.ResetCoordinates
Set Canvas = Page.Canvas

'bilhete
Set Param = Pdf.CreateParam("x=112; y=698; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("id"), Param, Font


'voucher
Set Param = Pdf.CreateParam("x=112; y=685; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("voucher"), Param, Font

'data emissao
Set Param = Pdf.CreateParam("x=474; y=698; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("regdata") , Param, Font


'data emissao
Set Param = Pdf.CreateParam("x=524; y=685; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("dataEmissao") , Param, Font

'pax
Set Param = Pdf.CreateParam("x=65; y=650; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText TRIM(voucherRS("nome"))& " " & TRIM(voucherRS("sobrenome")) , Param, Font

'documento
Set Param = Pdf.CreateParam("x=400; y=650; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText TRIM(voucherRS("documento")), Param, Font

'endereco
Set Param = Pdf.CreateParam("x=74; y=638; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("endereco") & " " & voucherRS("numero"), Param, Font

'bairro
Set Param = Pdf.CreateParam("x=341; y=638; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText trataNulo(voucherRS("bairro")), Param, Font

'cidade
Set Param = Pdf.CreateParam("x=65; y=626; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("cidade"), Param, Font

'uf
Set Param = Pdf.CreateParam("x=341; y=626; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("uf"), Param, Font

'cep
Set Param = Pdf.CreateParam("x=68; y=613; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText trataNulo(voucherRS("cep")), Param, Font

'dtNascimento
Set Param = Pdf.CreateParam("x=395; y=613; Size=9; color=black")
 Param("y") = Param("y") - 4
Canvas.DrawText voucherRS("dtNascimento"), Param, Font

if not isnull(voucherRS("beneficiario_nome")) then
	'beneficiario_nome
	Set Param = Pdf.CreateParam("x=80; y=599; Size=9; color=black")
	Canvas.DrawText trataNulo(voucherRS("beneficiario_nome")), Param, Font
end if

if not isnull(voucherRS("beneficiario_cpf")) then
	'beneficiario_cpf
	Set Param = Pdf.CreateParam("x=385; y=598; Size=9; color=black")
	Canvas.DrawText trataNulo(voucherRS("beneficiario_cpf")), Param, Font
end if



'plano
Set Param = Pdf.CreateParam("x=65; y=567; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText planoRS("nPlano") & "-" & planoRS("nome") , Param, Font

'destino
Set Param = Pdf.CreateParam("x=360; y=567; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText UCASE(destino)& "-"&voucherRS("destino") , Param, Font

'inicioVigencia
Set Param = Pdf.CreateParam("x=133; y=556; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText voucherRS("inicioVigencia"), Param, Font

'fimVigencia
Set Param = Pdf.CreateParam("x=422; y=556; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText voucherRS("fimVigencia"), Param, Font

'vigencia
Set Param = Pdf.CreateParam("x=530; y=556; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText voucherRS("dias"), Param, Font


''''''''' custos

premioUSD = formatNumber(voucherRS("valor_premio_USD"),2)
	
if  UCASE(planoRS("nacionalidade")) = "I" then
	premioBRL = formatNumber(premioUSD*voucherRS("cambio"),2)
else
	premioBRL = formatNumber(voucherRS("valor_premio_USD"),2)
end if
set addupRS = objConn.EXECUTE("SELECT COALESCE(SUM(valor_premio_USD),0) as premioUSD, COALESCE(SUM(valor_premio_USD),0) as premioBRL  FROM voucher where processoid='"&voucherRS("processoid")&"' and sobrenome='"&REPLACE(voucherRS("sobrenome"),"'","")&"' and voucher <> '"&voucherRS("voucher")&"' and flagUp = '1' and nome='"&voucherRS("nome")&"'")
if not addupRS.EOF then
	premioBRL = premioBRL + addupRS("premioBRL")
	if UCASE(planoRS("nacionalidade")) = "I" then
		premioUSD = premioUSD + addupRS("premioUSD")
	end if
	'	premioUSD = formatNumber(custoRS(0),2)
		
	if familiar = "1" then
		premioUSD = premioUSD/processoRS("nPax")
	end if
end if

''''''''''''''''''''

'tarifa USD
Set Param = Pdf.CreateParam("x=133; y=544; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText formatNumber(premioUSD,2), Param, Font


'cambio
Set Param = Pdf.CreateParam("x=347; y=544; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText formatNumber(voucherRS("cambio"),2), Param, Font



'liquido total a pagar
Set Param = Pdf.CreateParam("x=133; y=531; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText formatNumber(premioBRL-(premioBRL*(0.38/100)),2), Param, Font

'IOF
Set Param = Pdf.CreateParam("x=347; y=531; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText formatNumber( premioBRL*(0.38/100),2), Param, Font

'tarifa BRL
Set Param = Pdf.CreateParam("x=133; y=517; Size=9; color=black")
 Param("y") = Param("y") - 6
Canvas.DrawText formatNumber(premioBRL,2), Param, Font

' forma pgto
Set Param = Pdf.CreateParam("x=480; y=517; Size=9; color=black")
 Param("y") = Param("y") - 6
if voucherRS("pagamento") = "CC" then
pagamento = "Cart. de Cred."
else
pagamento = "a vista"
end if
Canvas.DrawText pagamento , Param, Font

' prazo pgto
Set Param = Pdf.CreateParam("x=548; y=517; Size=9; color=black")
 Param("y") = Param("y") - 6
	set prazoRS = objConn.execute("SELECT parcelas from emissaoprocesso where id = "&voucherRS("processoId"))
	parcelas = prazoRS("parcelas")
	set prazoRS = nothing
if parcelas = "" then parcelas = "1"
Canvas.DrawText parcelas&"x" , Param, Font



if not isnull(planoRS("OBS_SEG_IDADE")) and 1 <> 1 then 

Set Param = Pdf.CreateParam("x=32; y=114; Size=9; color=gray; width=390; height=200")
Canvas.DrawText planoRS("OBS_SEG_IDADE") , Param, Font

end if

'''' fator de cobertura inexistente
SQL = "  DECLARE @total as float, @n_cob as float, @dif as float, @fator as float " & _

"		 select  'total' = sum(pesohdi), " & _
"		 'n_cob' = count(id), " & _
"		 'dif' = 100 - sum(pesohdi), " & _
"		 'fator'= (100 - sum(pesohdi))  / count(id) " & _
		
"		 FROM coberturas where tipo = '1' and id in (SELECT coberturaid from coberturasplanos " & _

"		 								where    pesohdi <> 0 and ( " & _
"											coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"' ) "&sql_cob_upgrade&" order by nplano desc ) " & _
"									    	OR " & _
"											coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"')  "&sql_cob_upgrade&" ) " & _
"										)) "

'response.write SQL
'response.end()


set fatorRS = objConn.execute(SQL)
''''''''''''''''''''''''''''''''''''
'response.write SQL & "<BR><BR>"
'response.End()
if idioma = "1" then
	descritivo_cob = "descritivo"
	simbolo = "simbolo"
	valor_cob = "valor"
else
	descritivo_cob = "eng_descritivo"
	simbolo = "simbolo_ingles"
	valor_cob = "valor_ingles"

end if

if idioma = "3" then
	descritivo_cob = "esp_descritivo"
	simbolo = "simbolo_espanhol"
	valor_cob = "valor_espanhol"
end if


'''' montar cobertura de seguros
SQL = "  select " & _
"		 coberturas."&descritivo_cob&"  as descritivo, " & _
"		 coalesce(coberturas.pesoHdi,0) as pesoHdi, " & _
"		 coberturasplanos."&simbolo&"   as simbolo, " & _
"		 coberturasplanos."&valor_cob&" as valor " & _

"		FROM coberturas " & _

" 		inner join coberturasplanos ON coberturas.id = coberturasplanos.coberturaID " & _

"       where ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"'  )   )  " & _

"			    OR " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and coalesce(nPlano,0) = "&sql_cob_upgradeCov&"   order by nplano  )  " & _
"			    OR " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 0  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"'  )  "&sql_cob_upgrade&" ) " & _
"			  ) and tipo='1'  and coberturasplanos.versao_id = '"&versaoCob&"'  order by coberturas.ordem "

'response.Write SQL
'response.End()

set cobRS = objConn.execute(SQL)


Set param = Pdf.CreateParam("rows=17, cols=3, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 17
linha     = 0
conta_cob = 0

WHILE NOT cobRS.EOF and max_linha > linha

linha = linha + 1 
conta_cob = conta_cob + 1

if cobRS("pesoHDI") <> 0 then
	fator = fatorRS("fator")
else
	fator = 0
end if

	Table(linha,1).AddText TRIM(cobRS("descritivo")), "size=8; expand=true; color=black", Font
	Table(linha,1).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,1).Width   = "147"

	Table(linha,2).AddText TRIM(cobRS("simbolo"))&TRIM(cobRS("valor")), "size=8; expand=true; color=black", Font
	Table(linha,2).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,2).Width   = "74"
	
	Table(linha,3).AddText "R$" & formatNumber(premioBRL*((cobRS("pesoHDI")+fator)/100),4), "size=8; expand=true; color=black", Font
	Table(linha,3).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,3).Width   = "48"

cobRS.MOVENEXT
WEND

Canvas.DrawTable Table, "x=27, y=440"


Set param = Pdf.CreateParam("rows=17, cols=3, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 17
linha     = 0
'conta_cob = 0

WHILE NOT cobRS.EOF and max_linha > linha

linha = linha + 1 
conta_cob = conta_cob + 1

if cobRS("pesoHDI") <> 0 then
	fator = fatorRS("fator")
else
	fator = 0
end if

	Table(linha,1).AddText TRIM(cobRS("descritivo")), "size=8; expand=true; color=black", Font
	Table(linha,1).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,1).Width   = "147"

	Table(linha,2).AddText TRIM(cobRS("simbolo"))&TRIM(cobRS("valor")), "size=8; expand=true; color=black", Font
	Table(linha,2).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,2).Width   = "74"
	
	Table(linha,3).AddText "R$" & formatNumber(premioBRL*((cobRS("pesoHDI")+fator)/100),4), "size=8; expand=true; color=black", Font
	Table(linha,3).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	Table(linha,3).Width   = "48"

cobRS.MOVENEXT
WEND

Canvas.DrawTable Table, "x=299, y=440"

'' fim cob seguros

'condicoes gerais

ARQ_CG = LinkCG

Set Param = Pdf.CreateParam("x=107; y=62; Size=8; color=orange")
 Param("x") = Param("x") + 90
Canvas.DrawText "Cond. Gerais: " & ARQ_CG , Param, Font

''''' FIM PAGINA 2

'' pagina 3

' Obtain the only page's canvas
Set Page = Doc.Pages(3)
'resolve o problema das letras invertidas
Page.ResetCoordinates
Set Canvas = Page.Canvas
    
' QR CODE

       'Canvas.DrawBarcode2D "https://www.affinityseguro.com.br/condicoes-gerais/062019_CG.pdf", "Type=3; X=395; Y=738"
       '' Canvas.DrawBarcode2D "QR is scannable art", "Type=3; X=95; Y=738"

    qr_code = qrcodelink
			Set Image = Doc.OpenImage(qr_code)
			' Preserve aspect ratio
			'Width = 200
			'Height = Width * Image.Height / Image.Width
			'90 height and width
			Width = 65
            Height = 65
			
			ScaleX = Width / Image.Width * Image.ResolutionX / 72
			ScaleY = Height / Image.Height * Image.ResolutionY / 72
			
			Set Param = PDF.CreateParam
			Param("X") = 20
			Param("Y") = 780
			Param("ScaleX") = ScaleX
			Param("ScaleY") = ScaleY

	
		Canvas.DrawImage Image, Param
    
' Save document, the Save method returns generated file name
if idioma = "1" then
	Filename = Doc.Save(Server.MapPath("arquivos/"&LEFT(documento, 3)&id&".pdf"), True )
else
	Filename = Doc.Save(Server.MapPath("arquivos/EN"&LEFT(documento, 3)&id&".pdf"), True )
end if

voucherRS.close
objConn.close

set voucherRS = nothing
set objConn   = nothing

''''''' excluir arquivos antigos
Dim objFSO, strPath

strPath = Server.MapPath("arquivos")

Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Call Search (strPath)
 
Sub Search(str)
    Dim objFolder, objSubFolder, objFile, contador
    Set objFolder = objFSO.GetFolder(str)
	count = 0
	limit = 20

	response.write("pasta: "&str)

    For Each objFile In objFolder.Files    
        count = count + 1

		if count >= limit then
        	'objFile.Delete(True)
		end if        
    Next    
End Sub

if request.QueryString("mobile") = "S" then
	if idioma = "1" then
		response.Redirect "arquivos/"&LEFT(documento, 3)&id&".pdf"
	else
		response.Redirect "arquivos/EN"&LEFT(documento, 3)&id&".pdf"
	end if
end if

if request.QueryString("anexo") <> "1" then
	'é necessário passar o nome do arquivo no FORM
	Dim Arquivo
	
	if idioma = "1" then
		Arquivo = "arquivos/"&LEFT(documento, 3)&id&".pdf"
	else
		Arquivo = "arquivos/EN"&LEFT(documento, 3)&id&".pdf"
	end if
	 
	Response.Buffer = True
	Response.AddHeader "Content-Type","application/x-msdownload"
	Response.AddHeader "Content-Disposition","attachment; filename=" & Arquivo
	Response.Flush
	 
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
	objStream.Type = 1
	objStream.LoadFromFile Server.MapPath(Arquivo)
	Response.BinaryWrite objStream.Read
	objStream.Close
	Set objStream = Nothing
	Response.Flush
end if
%> 