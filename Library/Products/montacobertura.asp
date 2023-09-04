<%
function montaCobertura(id)

idv=id
textoVet="RESUMO DOS SERVI&Ccedil;OS DE ASSIST&Ecirc;NCIA EMERGENCIAL|LIMITES DE COBERTURAS|LIMITES DE INDENIZAÇ&Otilde;ES POR SEGURADO"
campo=""
obsLabel="Observações Gerais"

texto=split(textoVet,"|")

Dim objConf, cobertura, coberturaRS, nPRS2

Set  coberturaRS  =  objConn.Execute("select  *  from  coberturas WHERE tipo  ='1' ORDER BY ordem")
Set  nPRS  =  objConn.Execute("select  nPlano,publicado,ageId,novo from planos where id='"&idv&"'")

 
if not nPRS.EOF then
	if nPRS("publicado")="1" and nPRS("ageId") = "0" then
		Set  nPRS2  =  objConn.Execute("select  id, obs,eng_obs  from  planos  where id='"&idv&"'")
	end if
else

response.Write "(Erro: cobertura não encontrada.)"
response.End()
end if


Dim labServ, labLim, labValor

	labServ= texto(0)
	labLim= texto(1)
	labValor= texto(2)

	cobertura  =  "<table  width=  100%  border=0  cellspacing=0  cellpadding=2><tr><td  width=72%  bgcolor=#f78528><b><font  color=#FFFFFF >"&labServ&"</font></b></td><td  width=28%  bgcolor=#f78528><b><font  color=#FFFFFF>"&labLim&"</font></b></td></tr>"

WHILE NOT coberturaRS.EOF

Set  objConf  =  objConn.Execute("select  *  from  coberturasplanos  where coberturaId='"&coberturaRS("id")&"' AND versao_id = 2 AND planoId  ='"&nPRS2("id")&"'")


if not objConf.EOF then

		  cor1 = cor1 MOD 2
		  if cor1 = 0 then
		  BG1="#FFFFFF"
          else
          BG1="#E8EBEC"
          end if

descritivo=coberturaRS(campo&"descritivo")
if descritivo = "" or isnull(descritivo) then descritivo=coberturaRS("descritivo")

		cobertura  =  cobertura  &   "<tr>  <td  width=72%  bgcolor="&BG1&" >"&coberturaRS("id")&"-"&descritivo&"</td>  <td  width=28% bgcolor="&BG1&" >" &   objConf("simbolo")  & " " & objConf("valor") & "</td></tr>" 
				    cor1 = cor1 + 1
end if

coberturaRS.MOVENEXT
WEND
cobertura  =  cobertura  &  "</u></b></font></td></tr></table>"


montaCobertura = cobertura

end function
%>