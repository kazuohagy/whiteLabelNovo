<!--#include file="../Common/micMainCon.asp" -->
<!--#include file="../Common/funcoes.asp" -->
<!--#include file="../Common/cambio.asp" -->

<%  

	Session.CodePage=65001

Dim parcela

function achaTarifa_PUB(nPlano,vigencia,idade)
Dim planoRS, cotacaoCustoRS, tarifaPax, diferencaVigencia, erro

erro=0

'Seleciona o plano
Set tar_planoRS = objConn.Execute("SELECT * FROM planos where nPlano='"&nPlano&"' and COALESCE(ageId,0) = '0' and COALESCE(publicado,0)='1'")

if tar_planoRS.EOF then
	Set tar_planoRS = objConn.Execute("SELECT * FROM planos where nPlano='"&nPlano&"' and COALESCE(ageId,0) = '0' and COALESCE(operador,0)='1'")
end if
'if planoRS("publicado") = 0 then acordo=1


if erro=0 or 1=1 then

if tar_planoRS.EOF then
	Set tar_planoRS = objConn.Execute("SELECT top 1 * FROM planos where nPlano='"&nPlano&"' order by id")
end if
		
'verifica se o periodo da viagem ï¿½ menor q o limite
if (CINT(tar_planoRS("limiteAdicional")) >= CINT(vigencia)) then 

		Set cotacaoCustoRS = objConn.Execute("SELECT * FROM valoresdiarios where planoId='"&tar_planoRS("id")&"' and dias='"&vigencia&"'")
		'response.write ("SELECT * FROM valoresdiarios where planoId='"&tar_planoRS("id")&"' and dias='"&vigencia&"'")
		'response.End()
		if cotacaoCustoRS.EOF  then
			achaTarifa_PUB = "0"
		else
			achaTarifa_PUB = cotacaoCustoRS("preco")
		end if


		
		else
		
		'viagem maior q o valor maximo de custo por dia
		'calcula qts dias adicionais serao nescessarios para o calculo final dao valos
		diferencaVigencia = (vigencia - tar_planoRS("limiteAdicional")) 
			
		Set cotacaoCustoRS = objConn.Execute("SELECT * FROM valoresdiarios where planoId='"&tar_planoRS("id")&"' and dias='"&tar_planoRS("limiteAdicional")&"'")
		'soma os dias adicionais com o valor do custo do dia limite
		achaTarifa_PUB = (tar_planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco") 
			
end if

set verIdadeRS = objConn.execute("SELECT coalesce(idadeMinima50,0) as idade_Acrescimo, coalesce(idadeMinima,0) as idade_limite from planos where id = '"&tar_planoRS("id")&"'")

	if idade>verIdadeRS("idade_Acrescimo")  and verIdadeRS("idade_Acrescimo") <> 0 then
	
		achaTarifa_PUB = achaTarifa_PUB * 1.5
		
	end if

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
	set voucherRS      	= objConn.execute("SELECT * FROM voucher where voucher='"&request.QueryString("voucher")&"'")' and id='"&request.QueryString("id")&"' and coalesce(cancelado,0) = 0")
elseif request.QueryString("id") <> "" then
	set voucherRS      	= objConn.execute("SELECT * FROM voucher where id='"&request.QueryString("id")&"'")' and id='"&request.QueryString("id")&"' and coalesce(cancelado,0) = 0")

end if


if voucherRS.eof then
	response.write "Certificado nao localizado"
	response.End()
end if
    
if voucherRS("cancelado") = "1" then
	response.write "Certificado cancelado"
	response.End()
end if   

'''' tratar se e um plano de upgrade e redireciona para o principal

if voucherRS("flagUp") = "1" then

	set voucherARS = objConn.execute("SELECT voucher from voucher where coalesce(flagUp,0)=0 and processoId = '"&voucherRS("processoId")&"' and sobrenome='"&voucherRS("sobrenome")&"'")
	
	if not voucherARS.EOF then response.redirect "geraCertificado.asp?voucher="&voucherARS("voucher")

	set voucherARS = nothing

end if


'' verifica se tem upgrade
	sql_cob_upgrade = ""
	set voucherARS = objConn.execute("SELECT plano from voucher where coalesce(flagUp,0)=1 and processoId = '"&voucherRS("processoId")&"' and sobrenome='"&voucherRS("sobrenome")&"'")
	
	if not voucherARS.EOF then 
		while not voucherARS.eof 
			sql_cob_upgrade = sql_cob_upgrade & " OR nPlano = '"&voucherARS(0)&"' "
			voucherARS.movenext
		wend		
	end if			
	set voucherARS = nothing''''
    
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




if cdate(voucherRS("dataEmissao")) < cdate("01/03/2018") then

	response.Redirect "geraCertificado1.asp?" & request.QueryString


end if



                                                                                      
if cdate(voucherRS("dataEmissao")) <= cdate("31/12/2018") then

	response.Redirect "geraCertificadoSOMPO.asp?" & request.QueryString

end if

if voucherRS("clienteId") <> "15246" and voucherRS("clienteId") <> "19376" then

	'chubb affinity
	if cdate(voucherRS("dataEmissao")) >= cdate("15/04/2020") then
		certificado_modelo = "bilhete_chubb_2020_af.pdf"
		qrcodelink = Server.MapPath("qr-code/affinity/qr-af-2020.png")
	else
		certificado_modelo = "bilhete4_2018_AF.pdf"
		certificado_modelo = REPLACE(certificado_modelo,".pdf","_SANCOR.pdf")
		qrcodelink = Server.MapPath("qr-code/affinity/qr-af-2019.jpg")
	end if
	
else

	'chubb next
	if cdate(voucherRS("dataEmissao")) >= cdate("15/04/2020") then
		certificado_modelo = "bilhete_chubb_2020_nxt.pdf"
		LinkCG = "https://www.nextseguroviagem.com.br/skin/pdf/2020/CG_Chubb_Condicoes_Gerais.pdf"
		qrcodelink = Server.MapPath("qr-code/next/qr-next-2020.png")
	else
		certificado_modelo = "bilhete4_2018_AF_NEXT.pdf"
		certificado_modelo = REPLACE(certificado_modelo,".pdf","_SANCOR.pdf")
		LinkCG  = "https://www.nextseguroviagem.com.br/skin/pdf/2018/cg_viagem_2018.pdf"
		LinkCG  = "https://www.nextseguroviagem.com.br/skin/pdf/2019/cg_viagem_2019.pdf"
		qrcodelink = Server.MapPath("qr-code/next/qr-next-2019.png")
	end if 

end if

if voucherRS("clienteId") = "20163 " then
		if voucherRS("destino") = "17" then
		destino = "Estados Unidos + Mundo"
		end if
		if voucherRS("destino") = "6" then
		destino = "Europa"
		end if
		if voucherRS("destino") = "18" then
		destino = "America Latina"
		end if
		if voucherRS("destino") = "32" then
		destino = "Brasil"
		end if
		if voucherRS("destino") = "15" then
		destino = "Africa"
		end if
		if voucherRS("destino") = "19" then
		destino = "Asia" 
		end if
		if voucherRS("destino") = "21" then
		destino = "Oceania" 
		end if

else

	if IsNumeric(voucherRS("destino")) then

		set rsdetinoCombo = objConn.execute("select * from viagem_destino  where id='"&voucherRS("destino")&"'")
		destino = ucase(rsdetinoCombo("nome"))
	else
		destino = ucase(voucherRS("destino"))
	end if

end if

if idioma = "2" then certificado_modelo = REPLACE(certificado_modelo,".pdf","_EN.pdf")

if idioma = "3" then certificado_modelo = REPLACE(certificado_modelo,".pdf","_ES.pdf")



Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.OpenDocument( Server.MapPath("modelo/"&certificado_modelo) )
Set Font = Doc.Fonts("Arial")

' Obtain the only page's canvas
Set Canvas = Doc.Pages(1).Canvas


' logo da agencia
'verifica se esta agencia tem voucher personalizado -----------------------------
logo = ""
Set tempLogoRS = objConn.Execute("SELECT logo FROM cadCliente WHERE id = '"&voucherRS("clienteId")&"' ")

if tempLogoRS(0) = "1" then

	logo = Server.MapPath("../Images/agtaut/"&voucherRS("clienteId")&".jpg" )
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if objFSO.fileExists(logo) Then 
	
			Set Image = Doc.OpenImage(logo)
			' Preserve aspect ratio
			'Width = 200
			'Height = Width * Image.Height / Image.Width
			Width = 140
			Height = 40
			
			ScaleX = Width / Image.Width * Image.ResolutionX / 72
			ScaleY = Height / Image.Height * Image.ResolutionY / 72
			
			Set Param = PDF.CreateParam
			Param("X") = 190
			Param("Y") = 740
			Param("ScaleX") = ScaleX
			Param("ScaleY") = ScaleY

	
		Canvas.DrawImage Image, Param
	end if
	set objFSO = nothing
end if

	set tempLogoRS = nothing

'''' fim logo agencia



'voucher
Set Param = Pdf.CreateParam("x=112; y=712; Size=9; color=black")
Canvas.DrawText voucherRS("voucher"), Param, Font

'data emissao
Set Param = Pdf.CreateParam("x=475; y=712; Size=9; color=black")
Canvas.DrawText voucherRS("dataEmissao") , Param, Font



'pax
Set Param = Pdf.CreateParam("x=65; y=673; Size=9; color=black")
Canvas.DrawText TRIM(voucherRS("nome"))& " " & TRIM(voucherRS("sobrenome")) , Param, Font

'documento
Set Param = Pdf.CreateParam("x=385; y=673; Size=9; color=black")
Canvas.DrawText TRIM(voucherRS("documento")), Param, Font

'dtNascimento
Set Param = Pdf.CreateParam("x=525; y=673; Size=9; color=black")
Canvas.DrawText voucherRS("dtNascimento"), Param, Font

'endereco
Set Param = Pdf.CreateParam("x=74; y=661; Size=9; color=black")
Canvas.DrawText voucherRS("endereco") & " " & voucherRS("numero"), Param, Font

'bairro
Set Param = Pdf.CreateParam("x=341; y=661; Size=9; color=black")
Canvas.DrawText trataNulo(voucherRS("bairro")), Param, Font

'cidade
Set Param = Pdf.CreateParam("x=65; y=647; Size=9; color=black")
Canvas.DrawText voucherRS("cidade"), Param, Font

'uf
Set Param = Pdf.CreateParam("x=341; y=647; Size=9; color=black")
Canvas.DrawText voucherRS("uf"), Param, Font

'cep
Set Param = Pdf.CreateParam("x=500; y=647; Size=9; color=black")
Canvas.DrawText trataNulo(voucherRS("cep")), Param, Font

'plano
Set Param = Pdf.CreateParam("x=65; y=635; Size=9; color=black")
Canvas.DrawText planoRS("nPlano") & "-" & planoRS("nome") , Param, Font

'destino
Set Param = Pdf.CreateParam("x=360; y=635; Size=9; color=black")
Canvas.DrawText UCASE(destino) , Param, Font

'inicioVigencia
Set Param = Pdf.CreateParam("x=133; y=622; Size=9; color=black")
Canvas.DrawText CDATE(voucherRS("inicioVigencia")), Param, Font

'fimVigencia
Set Param = Pdf.CreateParam("x=420; y=622; Size=9; color=black")
Canvas.DrawText CDATE(voucherRS("fimVigencia")), Param, Font

'vigencia
Set Param = Pdf.CreateParam("x=530; y=622; Size=9; color=black")
Canvas.DrawText voucherRS("dias"), Param, Font

set addRS = objConn.EXECUTE("SELECT COALESCE(SUM(totalUSD),0) as totalUSD, COALESCE(SUM(totalBRL),0) as totalBRL  FROM voucher where processoid='"&voucherRS("processoid")&"' and sobrenome='"&REPLACE(voucherRS("sobrenome"),"'","")&"' and voucher <> '"&voucherRS("voucher")&"' and flagUp = '1'")
if not addRS.EOF then
	while not addRS.eof then
		premioBRL = premioBRL + addRS("totalBRL")
		if UCASE(planoRS("nacionalidade")) = "I" then
		premioUSD = premioUSD + addRS("totalUSD")
		end if
		addRS.movenext
	wend
end if

'tarifa USD
Set Param = Pdf.CreateParam("x=133; y=608; Size=9; color=black")
Canvas.DrawText formatNumber(premioUSD,2), Param, Font


'cambio
Set Param = Pdf.CreateParam("x=365; y=608; Size=9; color=black")
Canvas.DrawText formatNumber(voucherRS("cambio"),2), Param, Font



'tarifa BRL
Set Param = Pdf.CreateParam("x=480; y=608; Size=9; color=black")
Canvas.DrawText formatNumber(premioBRL,2), Param, Font

' forma pgto
Set Param = Pdf.CreateParam("x=133; y=595; Size=9; color=black")
if voucherRS("pagamento") = "CC" then
pagamento = "Cart. de Cred."
else
pagamento = "a vista"
end if
Canvas.DrawText pagamento , Param, Font

' prazo pgto
Set Param = Pdf.CreateParam("x=395; y=595; Size=9; color=black")
	set prazoRS = objConn.execute("SELECT parcelas from emissaoprocesso where id = "&voucherRS("processoId"))
	parcelas = prazoRS("parcelas")
	set prazoRS = nothing
if parcelas = "" then parcelas = "1"
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
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"' ) "&sql_cob_upgrade&" order by nplano ) )" & _
"				OR tipo='2' and coberturasplanos.versao_id = '"&versaoCob&"' and ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 0  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"' ) "&sql_cob_upgrade&" ) " & _
"			  )  order by coberturas.ordem "

'response.Write SQL
'response.End()

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

Canvas.DrawTable Table, "x=27, y=510"


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

Canvas.DrawTable Table, "x=299, y=510"


'servicos de seguro 

SQL = "  select " & _
"		 coberturas."&descritivo_cob&"  as descritivo, " & _
"		 coalesce(coberturas.pesoHdi,0) as pesoHdi, " & _
"		 coberturasplanos."&simbolo&"   as simbolo, " & _
"		 coberturasplanos."&valor_cob&" as valor " & _

"		FROM coberturas " & _

" 		inner join coberturasplanos ON coberturas.id = coberturasplanos.coberturaID " & _

"       where ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"'  )  "&sql_cob_upgrade&" order by nplano  )  " & _
"			    OR " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 0  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"'  )  "&sql_cob_upgrade&" ) " & _
"			  ) and tipo='1'  and coberturasplanos.versao_id = '"&versaoCob&"'  order by coberturas.ordem "

'response.Write SQL
'response.End()
set cobRS = objConn.execute(SQL)

Set param = Pdf.CreateParam("rows=15, cols=2, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 12
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



Canvas.DrawTable Table, "x=27, y=350"


Set param = Pdf.CreateParam("rows=15, cols=2, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 12
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
	'Table(linha,3).AddText "", "size=8; expand=true; color=black", Font
	'Table(linha,3).SetBorderParams  "Top=false; Bottom=false; Right=false; Left=false; BottomColor=gray"
	'Table(linha,3).Width   = "48"

cobRS.MOVENEXT
WEND

Canvas.DrawTable Table, "x=300, y=350"

'' fim cob servicos

'condicoes gerais

if voucherRS("dataEmissao") > CDATE("31/12/2017")  and  voucherRS("dataEmissao") < CDATE("15/04/2020") then 
	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_2018.pdf"
else
	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_2017.pdf"

end if
    
 if voucherRS("dataEmissao") >= CDATE("15/04/2020") then
    	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Chubb_Condicoes_Gerais.pdf"
	else
		ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/062019_CG.pdf"
end if
      


if idioma = "2" then 	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_EN.pdf"

if idioma = "3" then 	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_ES.pdf"


if voucherRS("clienteId") = "2855" then ARQ_CG = "https://seguroassistplus.com.br/wp-content/uploads/2017/08/Condiï¿½oes_Gerais_AssistPlus.pdf"

if voucherRS("clienteId") = "15246" or voucherRS("clienteId") = "19376" then ARQ_CG = LinkCG



'if idioma = "1" then
	Set Param = Pdf.CreateParam("x=33; y=67; Size=8; color=black")
'else
'	Set Param = Pdf.CreateParam("x=78; y=85; Size=8; color=blue")
'end if
Canvas.DrawText "Cond. Gerais: " & ARQ_CG , Param, Font

'' pagina 2
Set Canvas = Doc.Pages(2).Canvas

'bilhete
Set Param = Pdf.CreateParam("x=112; y=698; Size=9; color=black")
Canvas.DrawText voucherRS("id"), Param, Font

'if UCASE(planoRS("nacionalidade")) = "I" then
	'6916000004 ï¿½ Internacional
	convenio = "6916000007"
'else
	'6916000003 ï¿½ Nacional
'	convenio = "6916000002"
'end if
    
''   if voucherRS("clienteId") = "28" then  
    convenio = ""
 ''   end if
'convenio
Set Param = Pdf.CreateParam("x=450; y=722; Size=9; color=black")
Canvas.DrawText convenio , Param, Font

if voucherRS("dataEmissao") > CDATE("31/01/2017") then
	processoSusep = "15.414.900421/2016-13"
else
	if UCASE(planoRS("nacionalidade")) = "N" then
		processoSusep = "15414.901010/2015-64"
	else
		processoSusep = "15414.900989/2015-53"
	end if

end if

' processo
if voucherRS("dataEmissao") <= CDATE("31/12/2017") then
Set Param = Pdf.CreateParam("x=30; y=722; Size=10; color=black")
Canvas.DrawText "Processo SUSEP Nï¿½ " & processoSusep & " (Ramo 1369 - Viagem)" , Param, Font
end if


'voucher
Set Param = Pdf.CreateParam("x=112; y=685; Size=9; color=black")
Canvas.DrawText voucherRS("voucher"), Param, Font

'data emissao
Set Param = Pdf.CreateParam("x=524; y=698; Size=9; color=black")
Canvas.DrawText voucherRS("dataEmissao") , Param, Font


'data emissao
Set Param = Pdf.CreateParam("x=524; y=685; Size=9; color=black")
Canvas.DrawText voucherRS("dataEmissao") , Param, Font

'pax
Set Param = Pdf.CreateParam("x=65; y=647; Size=9; color=black")
Canvas.DrawText TRIM(voucherRS("nome"))& " " & TRIM(voucherRS("sobrenome")) , Param, Font

'documento
Set Param = Pdf.CreateParam("x=400; y=647; Size=9; color=black")
Canvas.DrawText TRIM(voucherRS("documento")), Param, Font

'endereco
Set Param = Pdf.CreateParam("x=74; y=635; Size=9; color=black")
Canvas.DrawText voucherRS("endereco") & " " & voucherRS("numero"), Param, Font

'bairro
Set Param = Pdf.CreateParam("x=341; y=635; Size=9; color=black")
Canvas.DrawText trataNulo(voucherRS("bairro")), Param, Font

'cidade
Set Param = Pdf.CreateParam("x=65; y=622; Size=9; color=black")
Canvas.DrawText voucherRS("cidade"), Param, Font

'uf
Set Param = Pdf.CreateParam("x=341; y=622; Size=9; color=black")
Canvas.DrawText voucherRS("uf"), Param, Font

'cep
Set Param = Pdf.CreateParam("x=68; y=611; Size=9; color=black")
Canvas.DrawText trataNulo(voucherRS("cep")), Param, Font

'dtNascimento
Set Param = Pdf.CreateParam("x=385; y=611; Size=9; color=black")
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
Canvas.DrawText planoRS("nPlano") & "-" & planoRS("nome") , Param, Font

'destino
Set Param = Pdf.CreateParam("x=360; y=567; Size=9; color=black")
Canvas.DrawText UCASE(destino)& "-"&voucherRS("destino") , Param, Font

'inicioVigencia
Set Param = Pdf.CreateParam("x=133; y=554; Size=9; color=black")
Canvas.DrawText voucherRS("inicioVigencia"), Param, Font

'fimVigencia
Set Param = Pdf.CreateParam("x=420; y=554; Size=9; color=black")
Canvas.DrawText voucherRS("fimVigencia"), Param, Font

'vigencia
Set Param = Pdf.CreateParam("x=530; y=554; Size=9; color=black")
Canvas.DrawText voucherRS("dias"), Param, Font



''''''''' custos
if voucherRS("dataEmissao") >= CDATE("15/04/2020") then
	premioUSD = formatNumber(voucherRS("valor_premio_USD"),2)
	
	if  UCASE(planoRS("nacionalidade")) = "I" then
		premioBRL = formatNumber(premioUSD*voucherRS("cambio"),2)
	else
		premioBRL = formatNumber(voucherRS("valor_premio_USD"),2)
	end if
else
	set custoRS = objConn.execute("SELECT round(custo,2) FROM custo_valor where idCusto = 3 and planoId = '"&voucherRS("planoid")&"' and dias = '"&voucherRS("dias")&"'")

	if custoRS.EOF then
		set custoRS = objConn.execute("SELECT round(custo,2) FROM custo_valor where idCusto = 3 and planoId IN (SELECT id from planos WHERE nPlano = '"&voucherRS("plano")&"') and dias = '"&voucherRS("dias")&"'")
	end if

	premioUSD = formatNumber(custoRS(0),2)
		
	if familiar = "1" then
		premioUSD = premioUSD/processoRS("nPax")
	end if
		
	if  UCASE(planoRS("nacionalidade")) = "I" then
		premioBRL = formatNumber(premioUSD*voucherRS("cambio"),2)
	else
		premioBRL = formatNumber(premioUSD,2)
	end if

	set custoRS = nothing
end if
''''''''''''''''''''

'tarifa USD
Set Param = Pdf.CreateParam("x=133; y=542; Size=9; color=black")
Canvas.DrawText formatNumber(premioUSD,2), Param, Font


'cambio
Set Param = Pdf.CreateParam("x=347; y=542; Size=9; color=black")
Canvas.DrawText formatNumber(voucherRS("cambio"),2), Param, Font

'liquido total a pagar
Set Param = Pdf.CreateParam("x=133; y=529; Size=9; color=black")
Canvas.DrawText formatNumber(premioBRL-(premioBRL*(0.38/100)),2), Param, Font

'IOF
Set Param = Pdf.CreateParam("x=347; y=529; Size=9; color=black")
Canvas.DrawText formatNumber(premioBRL*(0.38/100),2), Param, Font

'tarifa BRL
Set Param = Pdf.CreateParam("x=133; y=517; Size=9; color=black")
Canvas.DrawText formatNumber(premioBRL,2), Param, Font

' forma pgto
Set Param = Pdf.CreateParam("x=392; y=517; Size=9; color=black")
if voucherRS("pagamento") = "CC" then
pagamento = "Cart. de Cred."
else
pagamento = "a vista"
end if
Canvas.DrawText pagamento , Param, Font

' prazo pgto
Set Param = Pdf.CreateParam("x=548; y=517; Size=9; color=black")
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
"											coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and (nPlano ='"&plano_N&"' ) "&sql_cob_upgrade&" order by id desc ) " & _
"									    	OR " & _
"											coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"')  "&sql_cob_upgrade&" ) " & _
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


'''' montar cobertura de seguros
SQL = "  select " & _
"		 coberturas."&descritivo_cob&"  as descritivo, " & _
"		 coalesce(coberturas.pesoHdi,0) as pesoHdi, " & _
"		 coberturasplanos."&simbolo&"   as simbolo, " & _
"		 coberturasplanos."&valor_cob&" as valor " & _

"		FROM coberturas " & _

" 		inner join coberturasplanos ON coberturas.id = coberturasplanos.coberturaID " & _

"       where tipo='1' and coberturasplanos.versao_id = '"&versaoCob&"' AND ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 1  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)=0 and( nPlano ='"&plano_N&"'  ) "&sql_cob_upgrade&" order by nplano ) )" & _
"				OR tipo='1' and coberturasplanos.versao_id = '"&versaoCob&"' AND ( " & _
"				coberturasplanos.planoid IN (SELECT top 1 id from planos where  (coalesce(publicado,0) = 0  or coalesce(operador,0) = 1  or  familiar='1') and coalesce(ageID,0)='"&voucherRS("clienteId")&"' and (nPlano ='"&plano_N&"'  ) "&sql_cob_upgrade&"  ) " & _
"			  )  order by coberturas.ordem "

'response.Write SQL
'response.End()

set cobRS = objConn.execute(SQL)


Set param = Pdf.CreateParam("rows=15, cols=3, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 12
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


Set param = Pdf.CreateParam("rows=15, cols=3, width=274, height=181, CellBorder=0")
Set Table = Doc.CreateTable(param)
max_linha = 12
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

if voucherRS("dataEmissao") > CDATE("31/12/2017")  and  voucherRS("dataEmissao") < CDATE("15/04/2020") then 
	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_2018.pdf"
else
	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_2017.pdf"

end if
    
 if voucherRS("dataEmissao") >= CDATE("15/04/2020") then
    	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Chubb_Condicoes_Gerais.pdf"
	else
		ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/062019_CG.pdf"
end if

if idioma = "2" then 	ARQ_CG = "https://www.affinityseguro.com.br/condicoes-gerais/CG_Viagem_EN.pdf"



if voucherRS("clienteId") = "2855" then ARQ_CG = "https://seguroassistplus.com.br/wp-content/uploads/2017/08/Condiï¿½oes_Gerais_AssistPlus.pdf"

if voucherRS("clienteId") = "15246" or voucherRS("clienteId") = "19376" then ARQ_CG = LinkCG



'if idioma = "1" then
	Set Param = Pdf.CreateParam("x=33; y=67; Size=8; color=black")
'else
'	Set Param = Pdf.CreateParam("x=78; y=85; Size=8; color=blue")
'end if
Canvas.DrawText "Cond. Gerais: " & ARQ_CG , Param, Font

''''' FIM PAGINA 2


'' pagina 4

if certificado_modelo = "bilhete4_2018_AF.pdf" or certificado_modelo = "bilhete4_2018_AF_EN.pdf" or 1=1 then

' Obtain the only page's canvas
Set Canvas = Doc.Pages(4).Canvas
    
' QR CODE

       'Canvas.DrawBarcode2D "https://www.affinityseguro.com.br/condicoes-gerais/062019_CG.pdf", "Type=3; X=395; Y=738"
       '' Canvas.DrawBarcode2D "QR is scannable art", "Type=3; X=95; Y=738"

    qr_code = qrcodelink
			Set Image = Doc.OpenImage(qr_code)
			' Preserve aspect ratio
			'Width = 200
			'Height = Width * Image.Height / Image.Width
			Width = 90
            Height = 90
			
			ScaleX = Width / Image.Width * Image.ResolutionX / 72
			ScaleY = Height / Image.Height * Image.ResolutionY / 72
			
			Set Param = PDF.CreateParam
			Param("X") = 488
			Param("Y") = 748
			Param("ScaleX") = ScaleX
			Param("ScaleY") = ScaleY

	
		Canvas.DrawImage Image, Param
    
    
' bilhete
		Set Param = Pdf.CreateParam("x=95; y=738; Size=9; color=black")
		Canvas.DrawText voucherRS("id"), Param, Font
'voucher
		Set Param = Pdf.CreateParam("x=290; y=738; Size=9; color=black")
		Canvas.DrawText voucherRS("voucher"), Param, Font
'data de emissao
		Set Param = Pdf.CreateParam("x=495; y=738; Size=9; color=black")
		Canvas.DrawText voucherRS("dataEmissao"), Param, Font
end if


' Save document, the Save method returns generated file name
Filename = Doc.Save(Server.MapPath("/_certificados/arquivos/"&id&".pdf"), True )







'response.Redirect "arquivos/"&rcpArquivo&".pdf"
voucherRS.close
objConn.close

set voucherRS = nothing
set objConn   = nothing


''''''' excluir arquivos antigos
Dim objFSO, strPath

strPath = Server.MapPath("/_certificados/arquivos")


Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Call Search (strPath)
 
 
Sub Search(str)
    Dim objFolder, objSubFolder, objFile, contador
    contador = 0
    limite = 15
    Set objFolder = objFSO.GetFolder(str)
    For Each objFile In objFolder.Files
    
    if cint(contador) > cint(limite) then exit for
    
    	contador = contador + 1
 
        ' Use DateLastModified for modified date of a file
        If objFile.DateCreated < (Now() - 1) Then
            objFile.Delete(True)
        End If
 
    Next
    
    
End Sub

'''''''''''''''''''''''''

if request.QueryString("mobile") = "S" then

	response.Redirect "/_certificados/arquivos/"&id&".pdf"
	response.End()

end if

if request.QueryString("anexo") <> "1" then

	'ï¿½ necessï¿½rio passar o nome do arquivo no FORM
	Dim Arquivo
	Arquivo = "arquivos/"&id&".pdf"
	 
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