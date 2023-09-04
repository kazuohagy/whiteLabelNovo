<%
	Session.CodePage=65001

Dim  cobertura, confSQL, objConf, id, tipo, cor1, BG1, nPRS

Function RemoveAcentos(ByVal Texto)
		Dim ComAcentos
		Dim SemAcentos
		Dim Resultado
		Dim Cont
		'Conjunto de Caracteres com acentos
		ComAcentos = "ÁÍÓÚÉÄÏÖÜËÀÌÒÙÈÃÕÂÎÔÛÊáíóúéäïöüëàìòùèãõâîôûêÇç"
		'Conjunto de Caracteres sem acentos
		SemAcentos = "AIOUEAIOUEAIOUEAOAIOUEaioueaioueaioueaoaioueCc"
		Cont = 0
		Resultado = Texto
		Do While Cont < Len(ComAcentos)
			Cont = Cont + 1
			Resultado = Replace(Resultado, Mid(ComAcentos, Cont, 1), Mid(SemAcentos, Cont, 1))
		Loop
		RemoveAcentos = Resultado
End Function

function achaCobertura(cod)

	dim rstmp
	if isnull(cod) or cod="" or len(cod) <=0 then cod=0
	set rstmp = objConn.execute("SELECT descritivo FROM  coberturas WHERE id="&cod)
	if rstmp.eof then
		achaCobertura = "<b><font color='red'>Cobert. " & cod & ": N/A</font></b>"
	else
		achaCobertura = rstmp(0)
	end if
	rstmp.close
	set rstmp = nothing
end function

function formataIdioma(txt,idiom)
select case idiom
case 2:
formataIdioma = replace(txt,"SIM","YES")
formataIdioma = replace(formataIdioma,"NAO","NO")
formataIdioma = replace(formataIdioma,"NÃO","NO")
formataIdioma = replace(formataIdioma,"ATE","UNTIL")
formataIdioma = replace(formataIdioma,"ATÉ","UNTIL")
case else: formataIdioma=txt
end select
end function

function montaCobertura(id)
'idioma recebido na mesma variavel id, para evitar mudancas na estrutura da funcao:
varteste1=id
varteste2=replace(id,",","")
varteste3=replace(id,",","a")
if varteste1<>varteste2 and not(isnumeric(varteste3)) then idVet=split(id,",") : idv=idVet(0) : idioma=idVet(1) else idv=id
idv=replace(idv,",","")

'Seleção de idioma:
select case idioma
case 2:
'Inglês
textoVet="SUMMARY OF THE EMERGENCY ASSISTANCE SERVICES|COVERAGE LIMITS|LIMITS ON CLAIMS BY INSURED"
campo="eng_"
obsLabel="General Remarks"

case else:
'Português (idioma=1)
textoVet="RESUMO DOS SERVI&Ccedil;OS DE ASSIST&Ecirc;NCIA EMERGENCIAL|LIMITES DE COBERTURAS|LIMITES DE INDENIZA&Ccedil;&otilde;ES POR SEGURADO"
campo=""
obsLabel="Observa&ccedil;&otilde;es Gerais"
idioma = 1
end select

texto=split(textoVet,"|")

Dim objConf, cobertura, coberturaRS, nPRS2

Set  nPRS  =  objConn.Execute("select  nPlano,publicado,ageId,id  from  planos  where id='"&idv&"'")

if nPRS("publicado")="0" and nPRS("ageId") = "0" then
	Set  nPRS2  =  objConn.Execute("select  id, obs, eng_obs, nPlano  from  planos  where nPlano='"&nPRS(0)&"' AND ageId=0")
'	response.write 0
end if


if nPRS("publicado")="1" and nPRS("ageId") = "0" then
	Set  nPRS2  =  objConn.Execute("select  id, obs, eng_obs, nPlano  from  planos  where id='"&idv&"'")
'	response.write 1
end if

if nPRS("publicado")="0" and nPRS("ageId") <> "0" then
	Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano, eng_obs from  planos  where nPlano='"&nPRS(0)&"' AND ageId=0")
'	response.write 2
end if

if nPRS("publicado")="1" and nPRS("ageId") <> "0" then
	Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where nPlano='"&nPRS(0)&"' AND ageId='0'")
if nPRS2.EOF then Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where nPlano='"&nPRS(0)&"' AND ageId='0'")
'	response.write 3
end if 

if idv = 481 then Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where id='"&idv&"'")
if idv = 367 then Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where id='"&idv&"'")
if nPRS2.eof then Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where id='"&idv&"'")
Dim labServ, labLim, labValor


Set  coberturaRS  =  objConn.Execute("select  distinct coberturas.id as id, coberturas.descritivo as descritivo, coberturas.eng_descritivo as eng_descritivo, coberturasplanos.simbolo as simbolo, coberturasplanos.valor as valor, coberturas.ordem as ordem, coberturas.tipo as tipo  from  coberturas INNER JOIN coberturasplanos ON coberturas.id=coberturasPlanos.coberturaId WHERE coberturasPlanos.versao_id = '2' and   coberturasPlanos.planoId='"&nPRS2("id")&"' GROUP BY  coberturas.id, coberturas.descritivo, coberturas.eng_descritivo, coberturasplanos.simbolo, coberturasplanos.valor, coberturas.ordem, coberturas.tipo ORDER BY coberturas.tipo, coberturas.ordem")


	labServ= texto(0)
	labLim= texto(1)
	labValor= texto(2)

	cobertura  =  "<table  width='100%'  border=0  cellspacing=0  cellpadding=7><tr class='murcielago'style=' background-color:  #292d3c !important; border: 1px solid black;'><td  width='72%'  class='azul-claro'><b><font  color='#FFFFFF' >"&labServ&"</font></b></td><td  width='28%'  class='azul-claro' ><b><font  color='#FFFFFF'>"&labLim&"</font></b></td></tr>"

WHILE NOT coberturaRS.EOF

Set  objConf  =  objConn.Execute("select  *  from  coberturasplanos  where  coberturasPlanos.versao_id = '2' and coberturaId='"&coberturaRS("id")&"' AND planoId  ='"&nPRS2("id")&"'")



		  cor1 = cor1 MOD 2
		  if cor1 = 0 then
		  BG1="#FFFFFF"
          else
          BG1="#f4f7fc"
          end if

if idioma = 1 then descritivo=RemoveAcentos(coberturaRS("descritivo")) else descritivo=RemoveAcentos(coberturaRS("eng_descritivo"))
if descritivo = "" or isnull(descritivo) then descritivo=RemoveAcentos(coberturaRS("descritivo"))

cobertura  =  cobertura  &   "<tr style='border: 1px solid black;'>  <td  width='72%' height='25'  bgcolor="&BG1&" style='font-size:11px ' ><B>"& RemoveAcentos(ucase(descritivo)) &"</B></td>  <td  width='40%' bgcolor="&BG1&" style='font-size:11px'><B>" &  RemoveAcentos(objConf("simbolo"))  & " "
cor1 = cor1 + 1
cobertura  =  cobertura  & formataIdioma(objConf("valor"),idioma)& "</B> </td></tr>" 

		




coberturaRS.MOVENEXT
WEND
cobertura  =  cobertura  &  "</table>"

'response.Write(nPRS2(campo&"obs"))   '3995


	if nPRS2("obs") <> "" then
		if nPRS2(campo&"obs")="" or isnull(nPRS2(campo&"obs")) then 
			obsField=nPRS2("obs") 
		else 
		'response.Write(nPRS2(campo&"obs"))   '3995
			if nPRS2("id") = 3992 or nPRS2("id") = 3996 then
				
				obsField="Limite de dias : 120<br>Limite de idade:  Acima de 80 anos, o passageiro terá direito a 50% das coberturas."
			else
				obsField=nPRS2(campo&"obs")
			end if
		end if
		
		
		'cobertura  =  cobertura  &  "<table  width='100%'  border=0  cellspacing=0  cellpadding=2><tr  ><td class='azul-claro'><b><font  color='#FFFFFF' >"& RemoveAcentos(obsLabel) &"</font></b</td></tr><tr><td>"& RemoveAcentos(obsField) &"</td></tr>"
		
		cobertura  =  cobertura  &  "<table  width='100%'  border=0  cellspacing=0  cellpadding=2><tr style=' background-color:  #292d3c !important; '><td class='azul-claro' style='border: 1px solid black;'><b><font  color='#FFFFFF' >"& RemoveAcentos(obsLabel) &"</font></b></td></tr><tr><td  style='font-size:14px; background-color: #f4f7fc; border: 1px solid black;'>"& RemoveAcentos(obsField) &"</td></tr></table>"
	end if

montaCobertura = RemoveAcentos(cobertura)



coberturaRS.close
nPRS.close
nPRS2.close

set coberturaRS = nothing
set nPRS = nothing
set nPRS2 = nothing

end function


function montaCoberturaComp(vPlano, vDataInicio, vDataFim, vQtdDias, vNPax, vFamiliar)


	Dim vID(), vNomPlano(), vFlaFam(), vTemp, vCols, N , vCol
	dim vWHERE ,vWHERE2,  vCobertura , vOrd2
	
	vCols = 0
	vTemp = vPlano
	vRafa = vTemp
	vOrd2 = ""
	do while vTemp <> ""
		redim preserve vID(vCols + 1)
		redim preserve vNomPlano(vCols + 1)
		redim preserve vFlaFam(vCols + 1)
		if instr(vTemp, ";") > 0 then
			vID(vCols) = cint(mid(vTemp, 1, instr(vTemp, ";") -1))
			vTemp = mid(vTemp, instr(vTemp, ";")+1)
		else
			vID(vCols) = cint(vTemp)
			vTemp = ""
		end if
		vOrd2 = vOrd2 + "WHEN A.PlanoID = "&vID(vCols)&" THEN "&vCols&" "
		vCols = vCols + 1
	loop 
	if vCols = 0 then
		montaCoberturaComp = ""
		exit function
	end if 
	vWHERE = ""
	vWHERE2 = ""
	for N = 0 to vCols - 1
		vSQL = "SELECT A.id  FROM Planos A WHERE A.id = "&vID(N)&" "
		Set  coberturaRS  =  objConn.Execute(vSQL)
		vWHERE = vWHERE + "A.PlanoID = "&vID(N)&" OR "
		vWHERE2 = vWHERE2 + "A.ID = "&vID(N)&" OR "
		coberturaRS.close 
	next
	vWHERE = "(" + trim(mid(vWHERE, 1, len(vWHERE) - 3)) + ")"
	vWHERE2 = "(" + trim(mid(vWHERE2, 1, len(vWHERE2) - 3)) + ")"

	vSQL = ""
	vSQL = vSQL + "SELECT " 
	vSQL = vSQL + "   A.ID, "
	vSQL = vSQL + "   A.Nome,  "
	vSQL = vSQL + "   A.CodPlaFam,  "
	vSQL = vSQL + "   A.obs  "
	vSQL = vSQL + "FROM  "
	vSQL = vSQL + "   Planos A "
	vSQL = vSQL + "WHERE "
	vSQL = vSQL + "   " + vWHERE2 + " "
	Set  coberturaRS  =  objConn.Execute(vSQL)
	'rafael-5/6/09
	vRafa = split(vRafa,";")
	tamanhoArray=Ubound(vRafa)
	set coberturaRS2 = objConn.Execute("SELECT * FROM planos WHERE id='"&vRafa(0)&"'")
	obsFinal = "<b style='font:bold 12px verdana;' >"&coberturaRS2("nome")&"</b><br>"&coberturaRS2("obs")
	if tamanhoArray>=2 then
		set coberturaRS2 = objConn.Execute("SELECT * FROM planos WHERE id='"&vRafa(1)&"'")
		if not coberturaRS2.eof then obsFinal1 = "<br><br><b style='font:bold 12px verdana;' >"&coberturaRS2("nome")&"</b><br>"&coberturaRS2("obs")
	end if
	if tamanhoArray>=3 then
		set coberturaRS2 = objConn.Execute("SELECT * FROM planos WHERE id='"&vRafa(2)&"'")
		if not coberturaRS2.eof then obsFinal2 = "<br><br><b style='font:bold 12px verdana;' >"&coberturaRS2("nome")&"</b><br>"&coberturaRS2("obs")
	end if
	
	do until coberturaRS.eof 
		for n = 0 to vCols - 1
			if vID(n) = coberturaRS("ID") then
				if FunMostra(coberturaRS("CodPlaFam"), "") <> "" then
					vFlaFam(n) = 1
				else
					vFlaFam(n) = 0
				end if
				vNomPlano(n) = FunMostra(coberturaRS("Nome"), "")
			end if
		next
		coberturaRS.movenext 
	loop
	coberturaRS.close 
	'**************************************************************
	' Cabecalho
	'**************************************************************
	vCobertura = ""
	vCobertura = vCobertura + "<tr> " 
	vCobertura = vCobertura + "   <td>&nbsp;</td> "
	for n = 0 to vCols - 1
		if vCols = 1 then
			vCobertura = vCobertura + "   <th class=""planoTH_SEL"" > "
			vCobertura = vCobertura + "                           "&vNomPlano(n)&"  "
			vCobertura = vCobertura + "   </th> "
		else
			vCobertura = vCobertura + "   <th class=""planoTH"" > "
			vCobertura = vCobertura + "                           "&vNomPlano(n)&"  "
			vCobertura = vCobertura + "   </th> "
		end if
	next
	vCobertura = vCobertura + "</tr> " + chr(13) + chr(10)
	'**************************************************************
	' Valores
	'**************************************************************
	'vCobertura = vCobertura + "<tr> "
	'vCobertura = vCobertura + "   <td bgcolor='#E1E0E0' > "
	'vCobertura = vCobertura + "      <b><font color='#000000' ></font></b> "
	'vCobertura = vCobertura + "   </td>"
	'vCobertura = vCobertura + "   <td colspan='"&vCols&"' align ='center' bgcolor='#E1E0E0' > "
	'vCobertura = vCobertura + "       <b><font  color='#000000' >VALORES</font></b> "
	'vCobertura = vCobertura + "   </td> "
	'vCobertura = vCobertura + "</tr> " + chr(13) + chr(10)	
	vCotaTotal = ""
	vCotaPorPax = ""
	for n = 0 to vCols - 1  
		if n = 0 then
			vCor = "#EEBA4D"
		else
			vCor = "#F1F4Fc"
		end if	
		
		valorTotal = ""
		porPax  = ""
		
		'response.write "###" & vFamiliar & "###"
		if vFamiliar = 0 or vFamiliar = "" then
			cotar vID(n),vDataInicio,vDataFim,DateDiff("d",vDataInicio,vDataFim)+1,vNPax,0,0
		else
			cotar vID(n),vDataInicio,vDataFim,DateDiff("d",vDataInicio,vDataFim)+1,vNPax,0,vFamiliar
		end if
		if vID(n) = 4091 and vNPax = 2 then valorTotal = porPax
		
			
		vCotaTotal = vCotaTotal & "<TD align ='right' bgcolor='"&vCor&"' style='width:100px;' >" & valorTotal & "</td>"
		vCotaPorPax = vCotaPorPax & "<TD align ='right' bgcolor='"&vCor&"' >" & porPax & "</td>"
		vCotaPorPaxFam = vCotaPorPaxFam & "<TD align ='right' bgcolor='"&vCor&"' >" & porPaxFam & "</td>"
	next
	
	
	
	vCobertura = vCobertura + "<tr>" 
	vCobertura = vCobertura + "	   <td valign='top' bgcolor='#f4f7fe' >" 
	vCobertura = vCobertura + "	      <div align='right' ><b>Valor Por Passageiro </b>"
	vCobertura = vCobertura + "       </div>"
	vCobertura = vCobertura + "    </td>"
	vCobertura = vCobertura + vCotaPorPax
	vCobertura = vCobertura + "</tr>" + chr(13) + chr(10)
	
	
	'if vartype(vFamiliar)=8 THEN
		if vFamiliar = "" then
			vFamiliar = 0
		end if
	'end if
	if vFamiliar = 1 then
		vCobertura = vCobertura + "<tr>" 
		vCobertura = vCobertura + "	   <td valign='top' bgcolor='#f4f7fe' >" 
		vCobertura = vCobertura + "	      <div align='right' ><b>Valor por Acompanhante (até 4) </b>"
		vCobertura = vCobertura + "       </div>"
		vCobertura = vCobertura + "    </td>"
		vCobertura = vCobertura + vCotaPorPaxFam
		vCobertura = vCobertura + "</tr>" + chr(13) + chr(10)
	end if
	vCobertura = vCobertura + "<tr>"
	vCobertura = vCobertura + "	   <td valign='top' bgcolor='#f4f7fe' >" 
	vCobertura = vCobertura + "	      <div align='right' ><b>Valor Total</b>"
	vCobertura = vCobertura + "       </div>"
	vCobertura = vCobertura + "    </td>"
	vCobertura = vCobertura + vCotaTotal
	vCobertura = vCobertura + "</tr>" + chr(13) + chr(10)
	
	
		' botao comprar agora
	
	vCotaPorPax = ""
	for n = 0 to vCols - 1  
		if n = 0 then
			vCor = "#EEBA4D"
		else
			vCor = "#F1F4Fc"
		end if	
		
		valorTotal = ""
		porPax  = ""
		
		vComprarAgora = vComprarAgora & "<TD align ='center' bgcolor='"&vCor&"' ><div class=""comprarAgora""><a href=""../emissao/?planoid="&vID(n)&"&amp;dataInicio="&vDataInicio&"&amp;dataFim="&vDataFim&"&amp;destino="&vDestino&"&amp;nPax="&vNPax&""" target=""_parent"" >Comprar Agora</a></div></td>"
			
	next
	
	
	vCobertura = vCobertura + "<tr>" 
	vCobertura = vCobertura + "	   <td valign='top' bgcolor='#f4f7fe' >" 
	vCobertura = vCobertura + "	      <div align='right' >"
	vCobertura = vCobertura + "       </div>"
	vCobertura = vCobertura + "    </td>"
	vCobertura = vCobertura + vComprarAgora
	vCobertura = vCobertura + "</tr>" + chr(13) + chr(10)
	
	vCobertura = vCobertura + "<tr>" 
	vCobertura = vCobertura + "	   <td colspan='"&vCols+1&"' ><FONT color='red' >Preço sujeito a alteração devido a mudança do câmbio</FONT>"
	vCobertura = vCobertura + "	   </td> "
	vCobertura = vCobertura + "</tr>" + chr(13) + chr(10)
	
	'fim botao comprar agora

	'
	'**************************************************************
	' Beneficios e Servicos
	'**************************************************************
	vCobertura = vCobertura + "<tr> "
	vCobertura = vCobertura + "   <td bgcolor='#E1E0E0' > "
	vCobertura = vCobertura + "      <b><font color='#000000' >BENEFICIOS E SERVIÇOS</font></b> "
	vCobertura = vCobertura + "   </td> "
	vCobertura = vCobertura + "   <td colspan='"&vCols&"' align ='center' bgcolor='#E1E0E0' > "
	vCobertura = vCobertura + "       <b><font  color='#000000' >LIMITES DE COBERTURAS</font></b> "
	vCobertura = vCobertura + "   </td> "
	vCobertura = vCobertura + "</tr> " + chr(13) + chr(10)
	
	vSQL = ""
	vSQL = vSQL + "SELECT " 
	vSQL = vSQL + "   B.Ordem,  "
	vSQL = vSQL + "   A.coberturaId,  "
	vSQL = vSQL + "   CASE "
	vSQL = vSQL + "      " + vOrd2
	vSQL = vSQL + "   END AS OrdPla, "
	vSQL = vSQL + "   A.PlanoID,  "
	vSQL = vSQL + "   A.simbolo,  "
	vSQL = vSQL + "   A.valor,  "
	vSQL = vSQL + "   B.descritivo "
	vSQL = vSQL + "FROM  "
	vSQL = vSQL + "   coberturasplanos A "
	vSQL = vSQL + "   LEFT JOIN coberturas B "
	vSQL = vSQL + "      ON B.Id = A.coberturaId "
	vSQL = vSQL + "WHERE  A.versao_id = '2' and"
	vSQL = vSQL + "   " + vWHERE + " "
	vSQL = vSQL + "ORDER BY "
	vSQL = vSQL + "   B.ordem, "
	vSQL = vSQL + "   A.coberturaId, "
	vSQL = vSQL + "   CASE "
	vSQL = vSQL + "      " + vOrd2
	vSQL = vSQL + "   END  "
	'RESPONSE.Write(vSQL)
	Set  coberturaRS  =  objConn.Execute(vSQL)
	dim vOrdem , vCobID
	if coberturaRS.eof = false then
		vOrdem = 0
		vCobID = 0
		do Until coberturaRS.eof 
			if vOrdem <> coberturaRS("Ordem") or vCobID <> coberturaRS("coberturaId") then
				if vCobID <> 0 then
					for vCol = vCol to vCols - 1 
						if vCol = 0 then
							vCor = "#EEBA4D"
						else
							vCor = BG1
						end if	
						vCobertura = vCobertura & "<TD bgcolor='"&vCor&"' > &nbsp; </TD>"
					next
				end if
				if cor1 mod 2 = 0 then
					BG1="#FFFFFF"
				else
					BG1="#F1F4Fc"
				end if
				if vCobID = 0 then
					vCobertura = vCobertura & " </TR> " + chr(13) + chr(10)
				end if 
				vCobertura = vCobertura & "<TR>" + chr(13) + chr(10) + "<TD bgcolor='"&BG1&"' > " & coberturaRS("descritivo") & " </TD>" + chr(13) + chr(10) + ""
				vOrdem = coberturaRS("Ordem")
				vCobID = coberturaRS("coberturaId")
				vCol = 0
				cor1 = cor1 + 1
			end if
			for vCol = vCol to vCols - 1 
				if coberturaRS("PlanoID") = vID(vCol) then
					if vCol = 0 then
						vCor = "#EEBA4D"
					else
						vCor = BG1
					end if	
					if trim(FunMostra(coberturaRS("simbolo"), "")) <> "" then
						vCobertura = vCobertura & " <TD align ='right' bgcolor='"&vCor&"' > " & coberturaRS("simbolo") & coberturaRS("Valor") & " </TD>" + chr(13) + chr(10) + ""
					else
						vCobertura = vCobertura & " <TD align ='center' bgcolor='"&vCor&"' > " & coberturaRS("Valor") & " </TD>" + chr(13) + chr(10) + ""
					end if
					vCol = vCol + 1
					exit for
				else
					if vCol = 0 then
						vCor = "#EEBA4D"
					else
						vCor = BG1
					end if	
					vCobertura = vCobertura & " <td bgcolor='"&vCor&"' > &nbsp; </td>" + chr(13) + chr(10) + ""
				end if
			next
			coberturaRS.movenext
		loop 
	end if
	coberturaRS.close
	'vCobertura = vCobertura & "<TR bgcolor='#E1E0E0' ><TD colspan='2' ><b>Observações:</b></TD></TR>"
	'vCobertura = vCobertura & "<TR><TD colspan=4>"&obsFinal&obsFinal1&obsFinal2&"</TD></TR>"
	montaCoberturaComp = vCobertura
end function


%>