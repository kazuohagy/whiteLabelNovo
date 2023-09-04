<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/enviaEmail.asp" -->
<%

copia = Request.Form("copia")
destinatario = Request.Form("email")
voucher = Request.Form("voucher")
idioma = Request.Form("idioma")

vTemp = ""
vTemp = vTemp + "?voucher="&voucher
vTemp = vTemp + "&idioma="&request.Form("idioma")

set voucher_TempRs = objConn.execute("SELECT dataEmissao, id, voucher, email FROM voucher where voucher='"&voucher&"'")

k = voucher_TempRs("id")

Set objDynu = Server.Createobject("Dynu.HTTP") 

		objDynu.SetURL "https://www.affinityseguro.com.br/Library/Certificates/geraCertificado.asp?mobile=N&anexo=1&voucher="&voucher_TempRs("voucher")&"&idioma="&request.Form("idioma")
		objDynu.PostURL()
		'anexar arquivo PDF
		'voucherArquivo = replace(replace(replace(objRS("voucher"),".",""),"/",""),"-","")
		'voucherArquivo = objRS("url")
		pdfAnexoItem = Server.MapPath("../_certificados/arquivos/"&voucher_TempRs("id")&".pdf")
		HTML = HTML & "<BR>Segue anexo bilhete de viagem " & voucher_TempRs("voucher") & " (" & voucher_TempRs("id") & ".pdf)"
		
		
		pdfAnexo = pdfAnexo & "|" & pdfAnexoItem
		
		
		'''''''''''''''''''''' ENVIAR O EMAIL
		enviaMail "noreply@affinityseguro.com.br","Affinity Seguro","",destinatario,"Bilhete e Voucher Affinity Seguro - "&voucher_TempRs("voucher"),HTML,pdfAnexo

	if Request.Form("copia") <> "" then
		enviaMail "noreply@affinityseguro.com.br","Affinity Seguro","",copia,"Bilhete e Voucher Affinity Seguro - "&voucher_TempRs("voucher"),HTML,pdfAnexo
	end if

    objConn.Execute("INSERT INTO voucherEmail (destinatario, remetente , usuario,voucher, tipoEnvio) VALUES ('"&destinatario&"','"&remetente&"','"&request.cookies("FCNET_MIC")("login")&"','"&voucher&"','"&vt&"')")

%>