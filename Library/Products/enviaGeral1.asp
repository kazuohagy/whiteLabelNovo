<%
'-----------------------------------------------------
'Funcao: GerarTextoRandomico(ByVal TamanhoChave)
'Sinopse: Gerador de textos randômicos
'Parâmetros:
'	TamanhoChave: Optional (se vazio tamanho = 40)
'Retorno: String
'-----------------------------------------------------
function GerarTextoRandomico(TamanhoChave)
	Dim Chave
	Dim Num
	Dim arrValores(35)
	arrValores(0) 	=	"0"
	arrValores(1)	=	"1"
	arrValores(2)	=	"2"
	arrValores(3)	=	"3"
	arrValores(4)	=	"4"
	arrValores(5)	=	"5"
	arrValores(6)	=	"6"
	arrValores(7)	=	"7"
	arrValores(8)	=	"8"
	arrValores(9)	=	"9"
	arrValores(10)	=	"A"
	arrValores(11)	=	"B"
	arrValores(12)	=	"C"
	arrValores(13)	=	"D"
	arrValores(14)	=	"E"
	arrValores(15)	=	"F"
	arrValores(16)	=	"G"
	arrValores(17)	=	"H"
	arrValores(18)	=	"I"
	arrValores(19)	=	"J"
	arrValores(20)	=	"K"
	arrValores(21)	=	"L"
	arrValores(22)	=	"M"
	arrValores(23)	=	"N"
	arrValores(24)	=	"O"
	arrValores(25)	=	"P"
	arrValores(26)	=	"Q"
	arrValores(27)	=	"R"
	arrValores(28)	=	"S"
	arrValores(29)	=	"T"
	arrValores(30)	=	"U"
	arrValores(31)	=	"V"
	arrValores(32)	=	"W"
	arrValores(33)	=	"X"
	arrValores(34)	=	"Y"
	arrValores(35)	=	"Z"
	'Randomize em todo Array
	Randomize
	If TamanhoChave = "" Then
		TamanhoChave = 40
	End If
	Do While Len(Chave) < TamanhoChave
		Num = arrValores(Int(35 * Rnd ))
		Chave = Chave + Num
	Loop
	'Retornando a função
	GerarTextoRandomico = Chave
End Function


function enviaEvoucher(processo)
	
	set voucher_TempRs = objconn.execute("select * from voucher where processoId='"&processo&"' and coalesce(flagUp,0) = 0 order by id")
	voucher_id = voucher_TempRs("id")

	while not voucher_TempRs.eof
	
		if voucher_TempRs("flagUp") = "1" then
		
			nada = "nada"
		else

			pdfAnexo = ""
			HTML = ""
			documento = voucher_TempRs("documento")
			id = voucher_TempRs("id") 
			destinatario =  voucher_TempRs("email")
			
			HTML = HTML & "<BR>Segue o link para baixar seu bilhete de seguro viagem: <br><br><br><br> <a href='https://seguroviagemnext.com.br/Library/Certificates/geraCertificado.asp?mobile=N&voucher="&voucher_TempRs("voucher")&"&idioma=1' style='padding: 0.865rem;background-color: pink;font-weight:bolder;'>VOUCHER</a>"
				
			pdfAnexo = ""


			'''''''''''''''''''''' ENVIAR O EMAIL
			if voucher_TempRs("email") <> "" then
				enviaMail "","Next Seguro","",destinatario,"Voucher de seguro viagem - "&voucher_TempRs("voucher"),HTML,pdfAnexo
			end if

		

		'''''''''''''''''''''' gerar url do voucher para o SMS
		url = GerarTextoRandomico(3) & voucher_id & GerarTextoRandomico(3)
		
		objconn.execute("UPDATE voucher SET url = '"&url&"' where voucher = '"&voucher_TempRs("voucher")&"'")
		
		''''''''''''''''''''''''''
		
		pais = voucher_TempRs("pais")
		pais = 55 'deixando por enquanto, est�tico
		
		'''''''''''''''''''''' SMS INICIO
		
		para = pais & replace(replace(replace(replace(replace(voucher_TempRs("celular"),".",""),"-","")," ",""),")",""),"(","")
		
		if voucher_TempRs("clienteId") <> "2855" then

			msg = "Voucher "&voucher_TempRs("voucher")&" ("&voucher_TempRs("sobrenome")&"/"&voucher_TempRs("nome")&") "&vbcrlf&" Valido de "&voucher_TempRs("inicioVigencia")&" a "&voucher_TempRs("fimVigencia") & vbcrlf&" Acesse http://afy.me/"&url
			de = "Affinity"
		
		else
			msg = "Voucher ASSISTPLUS "&voucher_TempRs("voucher")&" ("&voucher_TempRs("sobrenome")&"/"&voucher_TempRs("nome")&") "&vbcrlf&"  Acesse http://afy.me/"&url
			de = "ASSISTPLUS"

		
		end if
		
		voucher = voucher_TempRs("voucher")

		
'		clienteId = 3202 
		 
'		parametros = "clienteId=3202"
'		parametros = parametros & "&chave=u210766816"
'		parametros = parametros & "&msg="&msg
'		parametros = parametros & "&to="&para
'		parametros = parametros & "&from=" & de
	
'		Set objDynu = Server.Createobject("Dynu.HTTP") 
'		objDynu.SetURL "https://ficopola.tavo.la/SMSService/enviar.asp"
'		objDynu.SetQueryString parametros
'		SMS = objDynu.PostURL()
'		'response.write SMS
'		Set objDynu = Nothing
		'''''''''''''''''''''' SMS FIM
		end if
		
	voucher_TempRs.movenext
	wend
	
	voucher_TempRs.close
	Set voucher_TempRs = Nothing 
	

end function
%>  