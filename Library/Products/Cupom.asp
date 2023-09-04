<%
	Dim novoValorBRL, novoValorUSD, novaComissaoBRL, novaComissaoUSD

	'retorna 
	function calculaDesconto (valorAntigo, percentual)

		if not isNumeric(valorAntigo) or not isNumeric(percentual) then
			response.write("Campos inválidos")
			response.end()
		end if 

		if percentual <= 0 then
			response.write("Não foi possível aplicar o desconto com os valores atuais")
			response.end()
		end if

		descontoZin = (100 - percentual) / 100
		calculaDesconto = Replace(valorAntigo, ".", ",") * descontoZin
		
	end function

	function calculaComissao ()
		
	end function 

    'metodo que insere no historico de cupom
	function insereCupomHistorico (validacao, cupomId, cupom, brl, usd, nBrl, nUsd)
	
		'codigo de validacao
		' 00 - cupom inválido
		' 01 - cupom valido
		
		historico = ""
			
		if validacao = "00" then
			historico = "cupom "& cupom &" invalido"
		end if
		
		if validacao = "01" then
			historico = "cupom "& cupom &" valido"
		end if
		
		if historico = "" then 
		
			response.write("Ocorreu um erro na verificação do cupom")
			
		else
	
			nBrl = replace(nBrl, ",", ".")
			nUsd = replace(nUsd, ",", ".")
			
			insertString = " INSERT INTO cupom_desconto_historico (cupom_id, processo_id, data, historico, usuario_login, totalBRLAnterior, totalBRLdesconto, totalUSDAnterior, totalUSDdesconto) " &_
				" VALUES ('"&cupomId&"', '"&processoId&"', GETDATE(), '"&historico&"' , '"&usuario&"' , "& brl &", "& nBrl &", "& usd &", "& nUsd &") "
		

			objConn.execute(insertString)
			
		end if
		
	end function

	'verifica se um cupom válido foi aplicado ao processo
	'retorna o percentual de desconto para o cálculo
	function getDesconto (cupomId, processoId, agencia, planoId)

		verificaQuery =    " SELECT c.idAge, cd.desconto " &_
							" FROM cupom_desconto cd " &_
	                    	" INNER JOIN comissao c ON cd.cadCliente_id = c.idAge " &_
							" INNER JOIN cupom_desconto_historico cdh ON cdh.processo_id = "& processoId &_
	                    	" WHERE cd.ativo = 1 " &_
                        	" AND cd.cadCliente_id = " & agencia &_
	                    	" AND cd.id = '"&cupomId&"' " &_
	                    	" AND c.idPlano = "& planoId &_
                        	" AND c.idAge = "& agencia &_
							" AND c.porcentagem >= cd.desconto "
		
	    set verificaRs = objConn.execute(verificaQuery)

		if not verificaRs.eof then
			getDesconto = verificaRs("desconto")
		end if
	
	end function



    'grava historico de cupom e retorna o desconto
	function aplicaCupom (cupom, processoId, agencia, planoId, brl, usd, totalBRL)

		'
	    cupomQuery =    " SELECT cd.id AS cpId, cd.desconto AS dsct, ep.desconto AS epCupomId, c.porcentagem " &_
                        " FROM cupom_desconto cd " &_
	                    " INNER JOIN comissao c ON cd.cadCliente_id = c.idAge " &_
	                    " INNER JOIN emissaoProcesso ep ON ep.id = "& processoId &_
	                    " WHERE cd.ativo = 1 " &_
                        " AND cd.cadCliente_id = " & agencia &_
	                    " AND cd.cupom = '"&cupom&"' " &_
	                    " AND c.idPlano = "& planoId &_
                        " AND c.idAge = "& agencia &_
						" AND c.porcentagem >= cd.desconto "&_
						" AND cd.data_expiracao >= GETDATE() "
						

	    set cupomRs = objConn.execute(cupomQuery)

	    'se cupom ok, calcula desconto e atualiza histórico
	    if not cupomRs.eof then 

			novoValorBRL = calculaDesconto(brl, cupomRs("dsct"))
			novoValorUSD = calculaDesconto(usd, cupomRs("dsct"))
			totalBRLdesconto = calculaDesconto(totalBRL, cupomRs("dsct"))
			totalComissao = Replace(totalBRL, ".", ",") * ((cupomRs("porcentagem")) / 100)
			novoTotalComissao = Replace(totalBRL, ".", ",") * ((cupomRs("porcentagem") - cupomRs("dsct")) / 100)

			'validacao
			if novoValorUSD <= 0 then
				insereCupomHistorico "00", "0", cupom, brl, usd, novoValorBRL, novoValorUSD
		    	response.ContentType = "application/json"
				response.write("{ ""erro"":""nao foi possivel adicionar desconto a esta compra"" }")
			end if 

			if cupomRs("epCupomId") <> "" and isNumeric(cupomRs("epCupomId")) then
				insereCupomHistorico "00", "0", cupom, brl, usd, novoValorBRL, novoValorUSD
		    	response.ContentType = "application/json"
				response.write("{ ""erro"":""um cupom ja foi adicionado a esta compra"" }")
			end if

			objConn.execute("UPDATE emissaoProcesso SET desconto = "& cupomRs("cpId") &" WHERE id ="& processoId)

		    insereCupomHistorico "01", cupomRS("cpId"), cupom, brl, usd, novoValorBRL, novoValorUSD

			response.ContentType = "application/json"

			response.write("{ 	""cupom"":"""& cupom &_	
								""",""antigoBRL"":"""& brl &_
								""",""antigoUSD"":"""& usd &_
								""",""novoBRL"":"""& novoValorBRL &_
								""", ""novoUSD"":"""& novoValorUSD &_
								""", ""percentual"":"""& cupomRs("dsct") &_
								""", ""comissaoPercent"":"""& cupomRs("porcentagem") &_
								""", ""totalBRLdesconto"":"""& totalBRLdesconto &_
								""", ""totalComissao"":"""& totalComissao &_
								""", ""novoTotalComissao"":"""& novoTotalComissao &_
							""" }")
	   
	    else
		    insereCupomHistorico "00", "0", cupom, brl, usd, 0, 0

			response.ContentType = "application/json"
			response.write("{ ""erro"":"""&cupom&""" }")
		    
	    end if
        
	end function 

	

%>