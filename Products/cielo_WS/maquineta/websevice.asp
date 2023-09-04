<%
b2w = ""
if processo = "" then processo = request.form("orderid")
if processo = "" then processo = request.QueryString("pedidoId")
if processo = "" then processo = request("processo")
if processo = "" then processo = 0

set testeHomolRS = objConn.execute("SELECT coalesce(b2w,0) FROM emissaoProcesso where id = '"&processo&"'")

if not testeHomolRS.EOF then
	if testeHomolRS(0) = 2 then
		'teste
		'afiliacaoCielo = "1006993069" 
		'chaveWS = "25fbb99741c739dd84d7b06ec78c9bac718838630f30b112d033ce2e621b34f3"
		'SoapServer = "https://qasecommerce.cielo.com.br/servicos/ecommwsec.do"
		b2w = "S"
		afiliacaoCielo = "1046652165" 
		chaveWS = "13fee37dd01e0a0057ad8a09a46276dd787cce5e48ec7eb9198c4f95c280df0a"
		SoapServer = "https://ecommerce.cielo.com.br/servicos/ecommwsec.do"
		
	else
		'producao
		afiliacaoCielo = "1046652165" 
		chaveWS = "13fee37dd01e0a0057ad8a09a46276dd787cce5e48ec7eb9198c4f95c280df0a"
		SoapServer = "https://ecommerce.cielo.com.br/servicos/ecommwsec.do"
	end if
else
	'producao
	afiliacaoCielo = "1046652165" 
	chaveWS = "13fee37dd01e0a0057ad8a09a46276dd787cce5e48ec7eb9198c4f95c280df0a"
	SoapServer = "https://ecommerce.cielo.com.br/servicos/ecommwsec.do"

end if


Function pegaValorNode(xml,node)

    Set objXml = Server.CreateObject("MSXML2.DOMDocument")

    objXml.loadXML(xml)

    If TypeName(objXml) = "DOMDocument" Then
        If objXml.GetElementsByTagName(node).length <> 0 Then
            pegaValorNode = objXml.selectSingleNode("//" & node).text
        Else
            pegaValorNode = ""
        End If
    Else
        pegaValorNode = ""
    End If

    Set objXml = Nothing

End Function

function statusCIELOTXT(codigo)

SELECT CASE codigo
	CASE "0":	statusCIELOTXT = "Criada"
	CASE "1":	statusCIELOTXT = "Em andamento"
	CASE "2":	statusCIELOTXT = "Autenticada"
	CASE "3":	statusCIELOTXT = "Não autenticada"
	CASE "4":	statusCIELOTXT = "Autorizada ou pendente de captura"
	CASE "5":	statusCIELOTXT = "Não autorizada"
	CASE "6":	statusCIELOTXT = "Capturada"
	CASE "8":	statusCIELOTXT = "Não capturada"
	CASE "9":	statusCIELOTXT = "Cancelada"
END SELECT



end function

%>