<%
Dim planoRS, cotacaoCustoRS, tarifaPax, porPax, valorTotal, tarifaPaxBRL, tarifaPaxUSD, diferencaVigencia,acordo, paxNUSD, paxNBRL, tt_valorParParcelar
Dim ttOriginalUSD, ttOriginalBRL
function cotar(planoId,dataInicio,dataFim,vigencia,nPax,destino,familiar)
		'response.write "FAMFGR" & familiar

'response.write "------------ " & familiar & " -----------------------"


if DateDiff("d",data(date,2,0),data(dataInicio,2,0)) < 0 then
	if request.QueryString("carregarReserva") = 1 then response.Redirect("../emissaoCliente/viagem_entra.asp?msg=1&"&request.QueryString)
	%><script>alert('\nA data de inicio da vigência é inferior a data atual.\nPor favor selecione outra data.')</script><%
	response.write "<script>window.history.back()</script>"
response.End()
end if


if DateDiff("d",data(date,2,0),data(dataFim,2,0)) < 0 then
	if request.QueryString("carregarReserva") = 1 then response.Redirect("../emissaoCliente/viagem_entra.asp?msg=2&"&request.QueryString)
	%><script>alert('\nA data de fim da vigência é inferior a data atual.\nPor favor selecione outra data.')</script><%
	response.write "<script>window.history.back()</script>"
response.End()
end if


if DateDiff("d",data(dataInicio,2,0),data(dataFim,2,0)) < 0 then
	if request.QueryString("carregarReserva") = 1 then response.Redirect("../emissaoCliente/viagem_entra.asp?msg=3&"&request.QueryString)
	%><script>alert('\nA data de fim da vigência é superior a data de inicio.\nPor favor selecione outra data.')</script><%
	response.write "<script>window.history.back()</script>"
response.End()
end if 

'Seleciona o plano

Set planoRS = objConn.Execute("SELECT * FROM planos where id='"&planoId&"'")
if planoRS("publicado") = 0 then acordo=1


if (planoId = "6255" or planoId = "6254") and (CDBL(nPax) < 2 or CDBL(nPax) > 10) then

	response.write "<script>alert('\Para categoria coletivo selecione de 2 a 10 passageiros.')</script>"
	response.write "<script>window.history.back()</script>"
	response.End()
end if



'verifica se pode ter adicional
if (planoRS("diaAdicional") = 0) AND (CINT(planoRS("limiteAdicional")) < CINT(vigencia)) then
	response.write "<script>alert('\nA vigência da viagem selecionada excedeu a vigência máxima deste plano.\nPor favor selecione outro plano.')</script>"
	response.write "<script>window.history.back()</script>"
response.End()
end if

'verifica o periodo maximo da emissao
if CINT(planoRS("vigenciaMaxima")) < CINT(vigencia) then
	response.write "<script>alert('\nA vigência da viagem selecionada excedeu a vigência máxima deste plano.\n\nVigência maxima para o plano " & planoRS("nome")&" é " & planoRS("vigenciaMaxima")&" dias. \n\nPor favor selecione outro plano ou outro periodo de viagem.')</script>"
	response.write "<script>window.history.back()</script>"
response.End()
end if

'verifica se é familiar e o numero de pax
if (familiar= "1") AND (nPax<2 OR npax > 5)  then
	response.write "<script>alert('\nPara obter desconto familiar é preciso ter pelo menos dois passageiros, e no máximo 5.')</script>"
	familiar = 0
	response.write "<script>window.history.back()</script>"
	response.End()
end if


		
'verifica se o periodo da viagem é menor q o limite
if (CINT(planoRS("limiteAdicional")) >= CINT(vigencia)) then

		Set cotacaoCustoRS = objConn.Execute("SELECT * FROM valoresdiarios where planoId='"&planoRS("id")&"' and dias='"&vigencia&"'")


		if cotacaoCustoRS.EOF  then
			response.write "<script>alert('\nNão existe tarifa cadastrada para o periodo de viagem selecionado\n\nPor favor selecione outro periodo.')</script>"
			response.write "<script>window.history.back()</script>"
		else
			if familiar <> "1" then			
				tarifaPax = (cotacaoCustoRS("preco"))
				if p50 = 1 then	tarifaPax = tarifaPax + (tarifaPax * 0.5)
			else
				'tarifaPax = (cotacaoCustoRS("precoFamiliar"))
				tarifaPax = 0
				if p50 = 1 then	tarifaPax = 0
			end if
		end if
		if cotacaoCustoRS.EOF then
		response.write "Nao ha tarifa para o periodo selecionado"
		response.End()
		end if
		ttOriginalUSD = cotacaoCustoRS("preco")
		if p50 = 1 then	ttOriginalUSD = ttOriginalUSD + (ttOriginalUSD * 0.5)
		
		if planoRS("nacionalidade") = "i" then
			ttOriginalBRL = cotacaoCustoRS("preco")* cambioTarifaRS("usdMic")
			if p50 = 1 then	ttOriginalBRL = ttOriginalBRL + (ttOriginalBRL * 0.5)
		else
			ttOriginalBRL = cotacaoCustoRS("preco")
			if p50 = 1 then	ttOriginalBRL = ttOriginalBRL + (ttOriginalBRL * 0.5)
		end if
		
		if session.SessionID = 792519043  then 
			response.Write("<br>--------------------------<br>")
			response.Write("cotacaoCustoRS = "&cotacaoCustoRS("preco")&"<br>")
			response.Write("ttOriginalUSD = "&ttOriginalUSD&"<br>")
			response.Write("cambioTarifaRS = "&cambioTarifaRS("usdMic")&"<br>")
			response.Write("ttOriginalBRL = "&ttOriginalBRL&"<br>")
			'response.End()
		end if
else
		
		'viagem maior q o valor maximo de custo por dia
		'calcula qts dias adicionais serao nescessarios para o calculo final dao valos
		diferencaVigencia = (vigencia - planoRS("limiteAdicional")) 
		Set cotacaoCustoRS = objConn.Execute("SELECT * FROM valoresdiarios where planoId='"&planoRS("id")&"' and dias='"&planoRS("limiteAdicional")&"'")
		
		'soma os dias adicionais com o valor do custo do dia limite	
		if familiar <> "1" then	
			tarifaPax = (planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco") 
			if p50 = 1 then	tarifaPax = tarifaPax + (tarifaPax * 0.5)
		else
			'tarifaPax = (planoRS("diaAdicionalFamiliar") * diferencaVigencia) + cotacaoCustoRS("precoFamiliar") 
			calc_x = (planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco") 
			tarifaPax = 0
			if p50 = 1 then	tarifaPax = 0
		end if	

		ttOriginalUSD = (planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco")
		if p50 = 1 then	ttOriginalUSD = ttOriginalUSD + (ttOriginalUSD * 0.5)
		if planoRS("nacionalidade") = "i" then
			ttOriginalBRL = formatNumber((planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco"),2) * formatNumber(cambioTarifaRS("usdMic"),2)
			if p50 = 1 then	ttOriginalBRL = ttOriginalBRL + (ttOriginalBRL * 0.5)
		else
			ttOriginalBRL = (planoRS("diaAdicional") * diferencaVigencia) + cotacaoCustoRS("preco")
			if p50 = 1 then	ttOriginalBRL = ttOriginalBRL + (ttOriginalBRL * 0.5)
		end if
end if  


		
		'response.Write(planoRS("diaAdicionalFamiliar"))
		'response.Write(" * ")
		'response.Write(diferencaVigencia)
		
		
		if planoRS("nacionalidade") = "i" then
			if familiar = "1" then
				if (CINT(planoRS("limiteAdicional")) >= CINT(vigencia)) then 'precisa verificar se o cálculo é feito com adicional ou não
					porPax = "US$ " & formatNumber(ttOriginalUSD) & "<BR>"
					tarifaPaxUSD = formatNumber(ttOriginalUSD)
					
					calcTT = ttOriginalUSD + (tarifaPax )'* (nPax_calc-1)
					cadcTTBRL = calcTT * cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "US$ " & formatNumber(calcTT,2)
					valorTotal = valorTotal & "<BR>Cambio: " & cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "<BR>R$ " & formatNumber(cadcTTBRL,2) &" <br>Primeiro passageiro tarifado, demais ZERADOS" 
					
					tt_valorParParcelar = cadcTTBRL
					tarifaPaxBRL = (formatNumber(tarifaPax)*formatNumber(cambioTarifaRS("usdMic")))
					paxNUSD = ttOriginalUSD
					paxNBRL = formatNumber(ttOriginalUSD)*formatNumber(cambioTarifaRS("usdMic"))
					tarifaPaxUSD = formatNumber(tarifaPax)
					
					if session.SessionID = 792519043  then 
					response.Write("<br><br>porPax = "&porPax&"<br>")
					response.Write("calcTT = "&calcTT&"<br>")
					response.Write("tarifaPaxUSD = "&tarifaPaxUSD&"<br>")
					response.Write("tarifaPax = "&tarifaPax&"<br>")
					response.Write(valorTotal&"<br>")
					'response.End()
					end if
					
				else
					porPax = "US$ " & formatNumber(calc_x) & "<BR>"
					tarifaPaxUSD = formatNumber(calc_x)
					
					calcTT = calc_x + (tarifaPax * (nPax_calc-1))
					cadcTTBRL = calcTT * cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "US$ " & formatNumber(calcTT,2)
					valorTotal = valorTotal & "<BR>Cambio: " & cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "<BR>R$ " & formatNumber(cadcTTBRL,2) &" <br>Primeiro passageiro tarifado, demais ZERADOS" 
					
					tt_valorParParcelar = cadcTTBRL
					tarifaPaxBRL = (formatNumber(tarifaPax)*formatNumber(cambioTarifaRS("usdMic")))
					paxNUSD = calc_x
					paxNBRL = formatNumber(calc_x)*formatNumber(cambioTarifaRS("usdMic"))
					tarifaPaxUSD = formatNumber(tarifaPax)
				end if
				

			else
				
				tarifaPaxUSD = formatNumber(tarifaPax)
				porPax = "US$" & formatNumber(tarifaPax)
				
				valorTotal = valorTotal & "US$ " & formatNumber((tarifaPax*nPax))
				valorTotal = valorTotal & "<BR>Cambio: " & cambioTarifaRS("usdMic")
				valorTotal = valorTotal & "<BR>R$ " & formatNumber((tarifaPax*nPax)*cambioTarifaRS("usdMic"))
				
				tarifaPaxBRL = (formatNumber(tarifaPax)*formatNumber(cambioTarifaRS("usdMic")))
			end if
				

		else
			if familiar = "1" then
				if (CINT(planoRS("limiteAdicional")) >= CINT(vigencia)) then 'precisa verificar se o cálculo é feito com adicional ou não
					porPax = "R$ " & formatNumber(cotacaoCustoRS("preco")) & "<BR>"
					tarifaPaxUSD = formatNumber(cotacaoCustoRS("preco"))

					calcTT = cotacaoCustoRS("preco") + (tarifaPax )'* (nPax_calc-1)
					cadcTTBRL = calcTT 
					valorTotal = valorTotal & "R$ " & formatNumber(calcTT,2)
					
					tt_valorParParcelar = cadcTTBRL
					tarifaPaxBRL = formatNumber(tarifaPax)
					paxNUSD = cotacaoCustoRS("preco")
					paxNBRL = formatNumber(cotacaoCustoRS("preco"))
					tarifaPaxUSD = formatNumber(tarifaPax)
				else
					porPax = "US$ " & formatNumber(calc_x) & "<BR>"
					tarifaPaxUSD = formatNumber(calc_x)
					
					calcTT = calc_x + (tarifaPax * (nPax_calc-1))
					cadcTTBRL = calcTT * cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "US$ " & formatNumber(calcTT,2)
					valorTotal = valorTotal & "<BR>Cambio: " & cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "<BR>R$ " & formatNumber(cadcTTBRL,2) 
					
					tt_valorParParcelar = cadcTTBRL
					tarifaPaxBRL = (formatNumber(tarifaPax)*formatNumber(cambioTarifaRS("usdMic")))
					paxNUSD = calc_x
					paxNBRL = formatNumber(calc_x)*formatNumber(cambioTarifaRS("usdMic"))
					tarifaPaxUSD = formatNumber(tarifaPax)
				end if
			
				'tarifaPaxUSD = formatNumber(0)
				'tarifaPaxBRL = formatNumber(tarifaPax)
				'porPax = "R$" & formatNumber(tarifaPax) & "<BR>(Primeiro Passageiro tarifa cheia, demais incluidos nesta tarifa)"
				'valorTotal = valorTotal & "<BR>R$ " & formatNumber(tarifaPax)
			else
				tarifaPaxUSD = formatNumber(0)
				tarifaPaxBRL = formatNumber(tarifaPax)
				porPax = "R$ " & formatNumber(tarifaPax)
				valorTotal = "R$ " & formatNumber(tarifaPax*nPax)
			end if
		end if		
		
			if planoId = 6255 or planoId = 6254 then
			valorTotal = ""
			porPax = "1o PAX" & porPax & "<br>DEMAIS 30% DE DESCONTO ATÉ 10 PAX" 
			
					calcTT = ttOriginalUSD + ((ttOriginalUSD-ttOriginalUSD*0.3)*(npax-1))'* (nPax_calc-1)
					cadcTTBRL = calcTT * cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "US$ " & formatNumber(calcTT,2)
					valorTotal = valorTotal & "<BR>Cambio: " & cambioTarifaRS("usdMic")
					valorTotal = valorTotal & "<BR>R$ " & formatNumber(cadcTTBRL,2) &" <br>Primeiro passageiro tarifado, demais com 30% de desconto!" 
					
					tt_valorParParcelar = cadcTTBRL
					tarifaPaxBRL = (formatNumber(tarifaPax)*formatNumber(cambioTarifaRS("usdMic")))
					paxNUSD = ttOriginalUSD
					paxNBRL = formatNumber(ttOriginalUSD)*formatNumber(cambioTarifaRS("usdMic"))
					tarifaPaxUSD = formatNumber(tarifaPax)

			
			
			end if
	

		
end function  

%>