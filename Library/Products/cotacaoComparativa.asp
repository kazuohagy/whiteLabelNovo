<%
function cotacaoComparativa(planos,vigencia,nPax,destino,familiar,cancelameto)

Dim strSQL, planoSQL, comparativoRS, vetorPlanos, i, corpo, planoLinha, descLina

planoLinha = "99999999"
descLina = "99999999"

vetorPlanos = SPLIT(planos,", ")

planoSQL = " coberturasPlanos.planoId="&vetorPlanos(0)

FOR i=1 to UBOUND(vetorPlanos)
	planoSQL = planoSQL & " OR coberturasPlanos.planoId="&vetorPlanos(i)
NEXT


strSQL = "select coberturas.Ordem as ordem, coberturas.Descritivo as descritivo, coberturasPlanos.simbolo as simbolo, coberturasPlanos.valor  as valor,  planos.familiar , planos.nome as plano, planos.nPlano as nPlano from coberturasPlanos INNER JOIN coberturas ON coberturasPlanos.coberturaId=coberturas.id INNER JOIN planos ON coberturasPlanos.planoId=planos.id WHERE coberturasPlanos.versao_id = '"&versao_cobertura&"' and ("

strSQL = strSQL & planoSQL

strSQL = strSQL & ") GROUP BY  planos.familiar , coberturas.Descritivo, coberturas.Ordem, coberturasPlanos.simbolo, coberturasPlanos.valor, planos.nome, planos.nPlano, planos.ordemExibicao ORDER BY ordem, Descritivo, ordemExibicao"


Set  comparativoRS  =  objConn.Execute(strSQL)

if cancelamento <> "" AND cancelamento <> "0000" then
	SQL = "SELECT planos.id, preco FROM planos_upgrade pU INNER JOIN valoresdiarios ON pU.upGradeId = valoresdiarios.planoId INNER JOIN planos ON pU.planoId = planos.id WHERE pU.upGradeId = "
	
		cancelSQL = cancelamento & " AND (planos.id = " & vetorPlanos(0)
		
		FOR i=1 to UBOUND(vetorPlanos)
			cancelSQL = cancelSQL & " OR planos.id = " & vetorPlanos(i)
		NEXT
						
	SQL = SQL & cancelSQL & ") ORDER BY ordemExibicao"
	
	set cancelRS = objConn.execute(SQL)
	
	if cancelRS.EOF then
		cancelamento = "0000"
	end if
end if

corpo = corpo & "<table  width=  100%  border=1  cellspacing=0  cellpadding=5 bgcolor=#EEEEEE>"

'inicio coberturas ==========================================

corpo = corpo & "<tr bgcolor=#C74801>"
corpo = corpo & "<td><b><font color=#ffffff>COBERTURAS</font></b></td>"

vetAqui = "#"
FOR i=0 to UBOUND(vetorPlanos)
	corpo = corpo & "<td bgcolor=#C74801><b><font color=#FFFFFF>"&achaPlano(vetorPlanos(i))&"</font></b></td>"
	vetAqui = vetAqui & "|"& achaPlano(vetorPlanos(i))
NEXT

vetAqui = replace(vetAqui,"#|","") 

coluna = 0
vet_aqui = split(vetAqui,"|")
dim nn
nn = 0
WHILE NOT comparativoRS.EOF

'response.write  nn & "<BR>"
'response.Flush()
	
	if descLina = "99999999" OR comparativoRS("descritivo") <> descLina then
		if descLina <> "99999999" then
		
			FOR i=0 to UBOUND(vetorPlanos)
				if vet_aqui(i)<>"" then
					corpo = corpo & vet_aqui(i)
				else
					corpo = corpo & "<td bgcolor=#FFFFFF><b>--</b></td>"
				end if
				vet_aqui(i) = ""
				nn = 0
			NEXT
		end if
		corpo = corpo & "</tr><tr>"
		corpo = corpo & "<td bgcolor=#FFFFFF>"&UCASE(comparativoRS("descritivo"))&"</td>"
		coluna=0
	end if
	
'	if achaPlano(vetorPlanos(coluna)) <> comparativoRS("plano") then
'		corpo = corpo & "<td bgcolor=#FFFFFF><b>--</b></td>"
'		coluna=coluna+1
'	end if

	if nn <= UBOUND(vet_aqui) then

		vet_aqui(nn) = "<td bgcolor=#FFFFFF>"&comparativoRS("simbolo") & " " & comparativoRS("valor") &"</td>"
	end if
	if nn < 4 then nn = nn +1
'	corpo = corpo & "<td bgcolor=#FFFFFF>"&comparativoRS("simbolo") & " " & comparativoRS("valor") &"</td>"
'	coluna=coluna+1

descLina = comparativoRS("descritivo")
planoLinha = comparativoRS("plano")

comparativoRS.MOVENEXt
WEND

FOR i=0 to UBOUND(vetorPlanos)
	if vet_aqui(i) <> "" then
		corpo = corpo & vet_aqui(i)
	else
		corpo = corpo & "<td bgcolor=#FFFFFF><b>--</b></td>"
	end if

NEXT
corpo = corpo&"</tr>"

'fim coberturas ==========================================
'inicio tarifas ==========================================
'TARIFA ORIGINAL =========================================
corpo = corpo & "<tr  bgcolor=#ffffff>"

if familiar="1"  then
	corpo = corpo & "<td><b>TARIFA TOTAL ("&nPax&" passageiro(s))</b></td>"
else
	corpo = corpo & "<td><b>TARIFA POR PASSAGEIRO US$</b></td>"
end if

FOR i=0 to UBOUND(vetorPlanos)

	
	set rsFPlano = objConn.execute("SELECT familiar FROM planos WHERE id = '"&vetorPlanos(i) &"'")
	
	if rsFPlano(0) = "1" then corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>US$ "&achaTarifaFam(vetorPlanos(i),destino,vigencia,"USD",1)&"</b></td>" else corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>US$ "&achaTarifa(vetorPlanos(i),destino,vigencia,"USD",1)&"</b></td>"

	
NEXT

corpo = corpo & "</tr>"

'CANCELAMENTO ==========================================
corpo = corpo & "<tr  bgcolor=#ffffff>"

corpo = corpo & "<td><b>UPGRADE DE CANCELAMENTO</b></td>"

if cancelamento <> "" AND cancelamento <> "0000" then
	FOR i=0 to UBOUND(vetorPlanos)	
		if NOT cancelRS.EOF then
			if CStr(cancelRS("id")) = vetorPlanos(i) then
				corpo = corpo & "<td ><b> US$ "& FormatNumber(cancelRS("preco"),2)&"</b></td>"
				cancelRS.MOVENEXT
			else
				corpo = corpo & "<td ><b>Não Disponível</b></td>"
			end if	
		else
			corpo = corpo & "<td ><b>Não Disponível</b></td>"
		end if		
	NEXT
	
	cancelRS.movefirst
else
	FOR i=0 to UBOUND(vetorPlanos)
		corpo = corpo & "<td ><b>Não Disponível</b></td>"
	NEXT		
end if	
corpo = corpo & "</tr>"

'CAMBIO ==========================================
corpo = corpo & "<tr  bgcolor=#ffffff>"

corpo = corpo & "<td><b>CAMBIO</b></td>"

FOR i=0 to UBOUND(vetorPlanos)
	if vetorPlanos(i) <> -1 then
	corpo = corpo & "<td ><b>"&cambioTarifaRS("usdMic")&"</b></td>"
	else
	corpo = corpo & "<td ><b>--</b></td>"
	end if
NEXT
corpo = corpo & "</tr>"

'TARIFA BRL ==========================================
corpo = corpo & "<tr  bgcolor=#ffffff>"


corpo = corpo & "<td><b>TARIFA POR PASSAGEIRO R$</b></td>"

FOR i=0 to UBOUND(vetorPlanos)
	set rsFPlano = objConn.execute("SELECT familiar FROM planos WHERE id = '"&vetorPlanos(i) &"'")

	if rsFPlano(0) = "1" then 
		corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "&achaTarifaFam(vetorPlanos(i),destino,vigencia,"BRL",1)&"</b></td>" 
	else 
		if cancelamento <> "" AND cancelamento <> "0000" then
			if NOT cancelRS.EOF then			
				if CStr(cancelRS("id")) = vetorPlanos(i)  then
					corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",1) + (cancelRS("preco")* cambioTarifaRS("usdMic"))&"</b></td>"
					cancelRS.MOVENEXT
				else
					corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",1) &"</b></td>"
				end if	
			else
				corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",1) &"</b></td>"
			end if
		else
			corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",1) &"</b></td>"
		end if
	end if
NEXT

if cancelamento <> "" AND cancelamento <> "0000" then 
	cancelRS.movefirst 
end if

corpo = corpo & "</tr>"

'TARIFA BRL TOTAL ==========================================
corpo = corpo & "<tr  bgcolor=#ffffff>"


corpo = corpo & "<td><b>TARIFA TOTAL R$ ("&nPax&" passageiro(s))</b></td>"

FOR i=0 to UBOUND(vetorPlanos)
	set rsFPlano = objConn.execute("SELECT familiar FROM planos WHERE id = '"&vetorPlanos(i) &"'")

	if rsFPlano(0) = "1" then 
		corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "&achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",1)&"</b></td>" 
	else 
		if cancelamento <> "" AND cancelamento <> "0000" then
			if NOT cancelRS.EOF then
				if CStr(cancelRS("id")) = vetorPlanos(i)  then
					corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL", nPax) + (cancelRS("preco") * cambioTarifaRS("usdMic") * nPax)&"</b></td>"
					cancelRS.MOVENEXT
				else
					corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",nPax) &"</b></td>"
				end if
			else
				corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",nPax) &"</b></td>"
			end if	
		else
			corpo = corpo & "<td style='border-bottom:#D9D9D9 1px solid'><b>R$ "& achaTarifa(vetorPlanos(i),destino,vigencia,"BRL",nPax) &"</b></td>"
		end if
	end if
NEXT

'fim tarifas ==========================================


'OBSERVAÇÕES GERAIS ==========================================
	strSQL = "select  distinct coberturasplanos.planoId , planos.nome, valoresDiarios.preco from coberturasPlanos INNER JOIN coberturas ON coberturasPlanos.coberturaId=coberturas.id INNER JOIN planos ON coberturasPlanos.planoId=planos.id LEFT JOIN valoresDiarios ON planos.id = valoresDiarios.planoId AND valoresDiarios.dias = 25 WHERE "
	strSQL = strSQL & planoSQL
	strSQL = strSQL & " ORDER BY valoresDiarios.preco, planos.nome "
	Set  rsListaPlanos  =  objConn.Execute(strSQL)

colunas = 1
while not rsListaPlanos.eof
colunas = colunas + 1 
rsListaPlanos.movenext
wend

corpo = corpo & "<tr  bgcolor=#C74801>"
corpo = corpo & "<td colspan='"&colunas&"'><font color=#FFFFFF><b>OBSERVAÇÕES GERAIS</b></FONT></td>"
corpo = corpo & "</tr>"

rsListaPlanos.movefirst

while not rsListaPlanos.eof
	Set  rsObs  =  objConn.Execute("select  id, obs, eng_obs, nPlano, nome  from  planos  where id='"&rsListaPlanos("planoId")&"'")
	corpo = corpo & "<tr bgcolor=#FFFFFF>"
	
	'if rsListaPlanos("planoId")= 6 then 
		'labelPlanoExcessao = "<FONT COLOR='RED'><B>(PRODUTO FEITO PARA AMERICA LATINA, NÃO RECOMENDADO PARA ESTADOS UNIDOS)<B></FONT>"
	'else
		labelPlanoExcessao = ""
	'end if
	corpo = corpo & "<td colspan='"&colunas&"' bgcolor=#ECF0F3><b><font color=#000000>"&rsObs("nome")&" "&labelPlanoExcessao&"</font></b></td>"
	corpo = corpo & "</tr>"
	
	corpo = corpo & "<tr bgcolor=#FFFFFF>"
	corpo = corpo & "<td colspan='"&colunas&"'><font color=#000000>"&rsObs("obs")&"</font></td>"
	corpo = corpo & "</tr>"
rsListaPlanos.movenext
wend

corpo = corpo & "</tr>"
corpo = corpo & "</table>" 


'if familiar="1" then corpo = corpo & "<small>Obs: A tarifa familiar indicada acima é válida para até 5 passageiros.<br>&nbsp;</small>"
cotacaoComparativa = corpo

end function
%>
