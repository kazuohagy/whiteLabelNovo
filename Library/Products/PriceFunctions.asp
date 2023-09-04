<%

set rateBRLRS = objConn.execute("select top 1 usdMic from cadCambio where data <= GETDATE() order by data desc")
rateBRL = rateBRLRS("usdMic")

Function CalcPremium (ByVal nPlano, ByVal days, ByVal destination, ByVal age, ByVal voucher, ByVal seqPro)
    if destination = "1" then 
        flagDestination = 1
    Else
        flagDestination = 0
    End if

    if age > 85 then
	    age = 85
    end if

    Set ageRS = objConn.Execute("SELECT id FROM idade WHERE '"& age &"' BETWEEN idadeMin AND idadeMax")
    if  ageRS.EOF then
        response.write "IDADE INEXISTENTE"
        response.End()
    end if   

    Set priceRS = objConn.Execute("SELECT premio, up_tipo_id FROM novo_custo_valor AS ncv INNER JOIN planos ON ncv.nPlano = planos.nPlano WHERE ncv.nPlano = '"& nPlano &"' and dias = '"& days &"' and flag_eua = '"& flagDestination &"'")    
    if  priceRS.EOF then
        Set priceRS = objConn.Execute("SELECT premio, up_tipo_id FROM novo_custo_valor AS ncv INNER JOIN planos ON ncv.nPlano = planos.nPlano WHERE ncv.nPlano = '"& nPlano &"' and up_tipo_id in (3,8) and flag_eua = '"& flagDestination &"'")
        if  priceRS.EOF then                        
            Set priceRS = objConn.Execute("SELECT cv.custo AS premio FROM custo_valor cv INNER JOIN planos ON cv.planoId = planos.id WHERE nPlano = '"& nPlano &"' AND dias = '"& days &"' AND idCusto = '1' AND (seqVar = 0 OR seqVar IS NULL )")
            if  priceRS.EOF then
                response.write "PREÇO INEXISTENTE"
                response.End()
            end if
        end if                            
    end if        
    
    premium = priceRS("premio")

    'covid
    If priceRS("up_tipo_id") = 8 Then    
        premium = premium * days * 0.75

    'cancelamento
	ElseIf priceRS("up_tipo_id") = 3 Then
		'não acontece nada

    'senior
    ElseIf priceRS("up_tipo_id") = 9 Then
        Set raiseRS = objConn.Execute("SELECT coeficiente FROM agravo_premio WHERE tipoAgravo = 1 and agravoId = '"& destination &"' or tipoAgravo = 2 and agravoId = 1")
        if  raiseRS.EOF then
            response.write "AGRAVO INEXISTENTE"
            response.End()
        end if

        If seqPro Mod 5 <> 1 then
            premium = premium - (premium * 0.25)
        End If

        WHILE NOT raiseRS.EOF
            premium = premium + (premium * raiseRS("coeficiente"))
        raiseRS.MOVENEXT
        WEND        

    'combo
    ElseIf priceRS("up_tipo_id") = 6 Then
        Set raiseRS = objConn.Execute("SELECT coeficiente FROM agravo_premio WHERE tipoAgravo = 1 and agravoId = '"& destination &"' or tipoAgravo = 2 and agravoId = '"& ageRS("id") &"'")
        if  raiseRS.EOF then
            response.write "AGRAVO INEXISTENTE"
            response.End()
        end if        

        if seqPro Mod 5 <> 1 then
            premium = premium - (premium * 0.25)
        End If

        WHILE NOT raiseRS.EOF
            premium = premium + (premium * raiseRS("coeficiente"))
        raiseRS.MOVENEXT
        WEND  

        Set upGradeRS = objConn.Execute("SELECT premio FROM novo_custo_valor WHERE nPlano = (SELECT nPlano FROM plano_combo WHERE comboPlano = "& nPlano &" and up_tipo_id = 8) and flag_eua = '"& flagDestination &"'")
        premium = premium + (upGradeRS("premio") * days * 0.75) 

    'covid+
    ElseIf priceRS("up_tipo_id") = 7 Then
        Set raiseRS = objConn.Execute("SELECT coeficiente FROM agravo_premio WHERE tipoAgravo = 1 and agravoId = '"& destination &"' or tipoAgravo = 2 and agravoId = '"& ageRS("id") &"'")
        if  raiseRS.EOF then
            response.write "AGRAVO INEXISTENTE"
            response.End()
        end if        

        if seqPro Mod 5 <> 1 then
            premium = premium - (premium * 0.25)
        End If

        WHILE NOT raiseRS.EOF
            premium = premium + (premium * raiseRS("coeficiente"))
        raiseRS.MOVENEXT
        WEND  

        Set upGradeRS = objConn.Execute("SELECT premio FROM novo_custo_valor WHERE nPlano = (SELECT nPlano FROM plano_combo WHERE comboPlano = "& nPlano &" and up_tipo_id = 8) and flag_eua = '"& flagDestination &"'")
        premium = premium + (upGradeRS("premio") * days)

        premium = premium + (1.48 * days)
             
    Else
        Set raiseRS = objConn.Execute("SELECT coeficiente FROM agravo_premio WHERE tipoAgravo = 1 and agravoId = '"& destination &"' or tipoAgravo = 2 and agravoId = '"& ageRS("id") &"'")
        if  raiseRS.EOF then
            response.write "AGRAVO INEXISTENTE"
            response.End()
        end if

        If (InStr(Right(voucher, 2),"F") <> 0 AND RIGHT(voucher, 2) <> "F1") OR (seqPro Mod 5 <> 1) then
            premium = premium - (premium * 0.25)
        End If

        WHILE NOT raiseRS.EOF
            premium = premium + (premium * raiseRS("coeficiente"))
        raiseRS.MOVENEXT
        WEND        
    end if  
	
	'Planos nacionais com 10% de desconto - nacional 3 e nacional 6
	if (nPlano = 583 or nPlano = 581) then
        premium = premium - (premium * 0.1)
    end if
	
	if nPlano = 886 then
        premium = premium - (premium * 0.2)
    end if

    if (nPlano = 542 or nPlano = 544 or nPlano = 887) AND days > 180 AND (destination = 5 OR destination = 8) then
        premium = premium + (premium * 0.2222)    
    end if        
    
	if (nPlano = 882) AND (destination = 5 OR destination = 8) then
        premium = premium + (premium * 0.1)    
    end if   	
	
	if priceRS("up_tipo_id") <> 3 AND priceRS("up_tipo_id") <> 8 AND nPlano <> 542 AND nPlano <> 544 AND nPlano <> 569 AND nPlano <> 570 AND nPlano <> 577 AND nPlano <> 887 AND nPlano <> 895 AND nPlano <> 896 AND nPlano <> 897 AND nPlano <> 882 then   
        if destination = 1 then        
            premium = premium + (premium * 0.4)
        else        
            premium = premium + (premium * 0.05)
        end if
    end if 
    
    CalcPremium = premium

End Function

Dim priceOriginal(), priceOriginalBR(), pricePax(), pricePaxBR(), priceTotal, priceTotalBR
Function CalcPrice (ByVal planoId, ByVal days, ByVal destination, ByVal age(), ByVal familiar)
    
    count = UBound(age)
    ReDim priceOriginal(count), priceOriginalBR(count), pricePax(count), pricePaxBR(count)
    priceTotal = 0
    priceTotalBR = 0

    Set planRS = objConn.Execute("SELECT limiteAdicional, diaAdicional, nacionalidade, nPlano FROM planos where id='"& planoId &"'")
    dayMax = planRS("limiteAdicional")

    CalcPrice = "Cambio: " & rateBRL
    CalcPrice = CalcPrice & "<br>LimiteAdicional: " & dayMax
    CalcPrice = CalcPrice & "<br>Dias: " & days

    if cint(days) > cint(dayMax) then
        Set priceRS = objConn.Execute("SELECT preco FROM valoresdiarios where planoId='"& planoId &"' and dias='"& dayMax &"'")
        
        addDays = days - dayMax
        price = (planRS("diaAdicional") * addDays) + priceRS("preco")
        CalcPrice = CalcPrice & "<br><br>Preço com dias adicionais: $" & price
    else
        Set priceRS = objConn.Execute("SELECT preco FROM valoresdiarios where planoId= '"& planoId &"' and dias= '"& days &"'")

        price = priceRS("preco")        
        CalcPrice = CalcPrice & "<br><br>Preço normal: $" & price
    end if

    Set raiseRS = objConn.Execute("SELECT SUM(coeficiente) as raise FROM agravo_venda WHERE plano = '"& planoId &"' and destino = '"& destination &"'")
    if IsNull(raiseRS("raise")) then
        raise = 0
        CalcPrice = CalcPrice & "<br>Desconto não foi encontrado: 0<br>"
    else
        raise = raiseRS("raise")
        CalcPrice = CalcPrice & "<br>Desconto é: " & raise & "<br>"
    end if

    For i = 0 To count
        priceTemp = price

        if age(i) >= 65 then
            priceTemp = priceTemp * 1.5
            CalcPrice = CalcPrice & "<br>Preço com Agravo de Idade: $" & priceTemp
        end if

        priceOriginal(i) = priceTemp
        priceOriginalBR(i) = priceTemp * rateBRL

        if (planoId = 6255 or planoId = 6254) and i > 0 then
            priceTemp = priceTemp - (priceTemp * 0.3)
            CalcPrice = CalcPrice & "<br>Preço com Coletivo: $" & priceTemp
        end if

        if familiar = "1" and i > 0 then
            priceTemp = 0
            CalcPrice = CalcPrice & "<br>Preço com familiar: $0"
        end if

        priceTemp = priceTemp + (priceTemp * raise)

        pricePax(i) = priceTemp
        pricePaxBR(i) = priceTemp * rateBRL
        priceTotal = priceTotal + pricePax(i)
        priceTotalBR = priceTotalBR + pricePaxBR(i)

        if planRS("nacionalidade") = "n" then
            priceOriginalBR(i) = NULL
            pricePaxBR(i) = NULL
            priceTotalBR = NULL
            CalcPrice = CalcPrice & "<br>Plano é Nacional"
        end if

        CalcPrice = CalcPrice & "<br>Preço Original: $" & priceOriginal(i) & " / R$" & priceOriginalBR(i) & "<br>Preç o por Pax: $" & pricePax(i) & " / R$" & pricePaxBR(i) & "<br>"
    Next

    CalcPrice = CalcPrice & "<br><br>Preço Total: $" & priceTotal & " / R$" & priceTotalBR

End Function

Dim priceCovid, priceCovidOriginal
Function CalcCovid (ByVal covidId, ByVal planoId, ByVal destination, ByVal age, ByVal days)
    
    set covidRS = objConn.execute("SELECT preco FROM planos_upgrade inner join valoresdiarios on upGradeId = valoresdiarios.planoId WHERE planos_upgrade.planoId = '"& planoId &"' and upGradeId = '"& covidId &"'")					

    Set raiseRS = objConn.Execute("SELECT SUM(coeficiente) as raise FROM agravo_venda WHERE plano = '"& covidId &"' and destino = '"& destination &"'")

    if IsNull(raiseRS("raise")) then
        raise = 0        
    else
        raise = raiseRS("raise")        
    end if

    if not covidRS.eof then	
        priceCovidOriginal = covidRS("preco") * days
        priceCovid = priceCovidOriginal + (priceCovidOriginal * raise)
    else
        priceCovidOriginal = 0
        priceCovid = 0        
    end if

End Function

Dim priceCancel
Function CalcCancel (ByVal cancelId, ByVal planoId)
    
    set cancelRS = objConn.execute("SELECT preco FROM planos_upgrade inner join valoresdiarios on upGradeId = valoresdiarios.planoId WHERE planos_upgrade.planoId = '"& planoId &"' and upGradeId = '"& cancelId &"'")					

    if not cancelRS.eof then	
        priceCancel = cancelRS("preco")
    else
        priceCancel = 0        
    end if

End Function

Function CalcPriceCambio (ByVal planoId, ByVal days, ByVal destination, ByVal age(), ByVal familiar, ByVal cambioZin)

	if cambioZin = "" or cambioZin = 0 then
		CalcPriceCambio = CalcPriceCambio & "<br>Câmbio não informado<br>" 
	end if
    
    count = UBound(age)
    ReDim priceOriginal(count), priceOriginalBR(count), pricePax(count), pricePaxBR(count)
    priceTotal = 0
    priceTotalBR = 0

    Set planRS = objConn.Execute("SELECT limiteAdicional, diaAdicional, nacionalidade, nPlano FROM planos where id='"& planoId &"'")
    dayMax = planRS("limiteAdicional")

    CalcPriceCambio = "Cambio: " & cambioZin
    CalcPriceCambio = CalcPriceCambio & "<br>LimiteAdicional: " & dayMax
    CalcPriceCambio = CalcPriceCambio & "<br>Dias: " & days

    if cint(days) > cint(dayMax) then
        Set priceRS = objConn.Execute("SELECT preco FROM valoresdiarios where planoId='"& planoId &"' and dias='"& dayMax &"'")
        
        addDays = days - dayMax
        price = (planRS("diaAdicional") * addDays) + priceRS("preco")
        CalcPriceCambio = CalcPriceCambio & "<br><br>Preço com dias adicionais: $" & price
    else
        Set priceRS = objConn.Execute("SELECT preco FROM valoresdiarios where planoId= '"& planoId &"' and dias= '"& days &"'")

        price = priceRS("preco")        
        CalcPriceCambio = CalcPriceCambio & "<br><br>Preço normal: $" & price
    end if

    Set raiseRS = objConn.Execute("SELECT SUM(coeficiente) as raise FROM agravo_venda WHERE plano = '"& planoId &"' and destino = '"& destination &"'")
    if IsNull(raiseRS("raise")) then
        raise = 0
        CalcPriceCambio = CalcPriceCambio & "<br>Desconto não foi encontrado: 0<br>" 
    else
        raise = raiseRS("raise")
        CalcPriceCambio = CalcPriceCambio & "<br>Desconto é: " & raise & "<br>"
    end if

    For i = 0 To count
        priceTemp = price

        if age(i) >= 65 then
            priceTemp = priceTemp * 1.5
            CalcPriceCambio = CalcPriceCambio & "<br>Preço com Agravo de Idade: $" & priceTemp
        end if

        priceOriginal(i) = priceTemp
        priceOriginalBR(i) = priceTemp * cambioZin

        if (planoId = 6255 or planoId = 6254) and i > 0 then
            priceTemp = priceTemp - (priceTemp * 0.3)
            CalcPriceCambio = CalcPriceCambio & "<br>Preço com Coletivo: $" & priceTemp
        end if

        if familiar = "1" and i > 0 then
            priceTemp = 0
            CalcPriceCambio = CalcPriceCambio & "<br>Preço com familiar: $0"
        end if

        priceTemp = priceTemp + (priceTemp * raise)

        pricePax(i) = priceTemp
        pricePaxBR(i) = priceTemp * cambioZin
        priceTotal = priceTotal + pricePax(i)
        priceTotalBR = priceTotalBR + pricePaxBR(i)

        if planRS("nacionalidade") = "n" then
            priceOriginalBR(i) = NULL
            pricePaxBR(i) = NULL
            priceTotalBR = NULL
            CalcPriceCambio = CalcPriceCambio & "<br>Plano é Nacional"
        end if

        CalcPriceCambio = CalcPriceCambio & "<br>Preço Original: $" & priceOriginal(i) & " / R$" & priceOriginalBR(i) & "<br>Preç o por Pax: $" & pricePax(i) & " / R$" & pricePaxBR(i) & "<br>"
    Next

    CalcPriceCambio = CalcPriceCambio & "<br><br>Preço Total: $" & priceTotal & " / R$" & priceTotalBR

End Function
%>