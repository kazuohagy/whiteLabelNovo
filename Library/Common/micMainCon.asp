<%
	Session.CodePage=65001
Response.AddHeader "P3P", "CP='NOI CUR OUR IND UNI COM NAV INT"'"

'Option Explicit
'Session.LCID = 1046
'Response.Expires = -1
'Response.ExpiresAbsolute = Now() - 2
'Response.AddHeader "pragma","no-cache"
'Response.AddHeader "cache-control","private"
'Response.CacheControl = "No-Store"

Dim objConn

COMPANY_NAME = "Next Seguro Viagem"
COMPANY_URL = "www.nextseguro.com.br"
COMPANY_WL_URL = "localhost"
COMPANY_LOGO = "../Images/logo.png"

Set objConn =  Server.CreateObject("ADODB.Connection")
velho = DATE()
novo = #23/08/2023#
if velho < novo then
	objConn.Open "DRIVER={SQL Server};SERVER=9.0.0.22;DATABASE=next;UID=next;PWD=Next@2021!@#"
else
	objConn.Open "DRIVER={SQL Server};SERVER=9.0.0.175;DATABASE=next;UID=next;PWD=Next@2021!@#"

end if 

function verLogado(logado,endereco)
	if logado<>"1"  then
		response.cookies("ficopolaCA")("irpara") = endereco
		response.redirect "../login/index.asp?irpara="&endereco
	end if
end function


function protetorSQL(campo)

'remove palavras que contenham sintaxe sql
termos_proibidos =  "like '%|like| or | from |select|insert|delete|update |where|drop table|show tables|#|\*|'|char)|dbo."

vertorTermos = SPLIT(termos_proibidos,"|")

FOR i=0 to UBOUND(vertorTermos)



	achaTermo = InStr(campo,vertorTermos(i))		
	if achaTermo <> 0 then
	resultado = "Invalid Input " & campo & "  " & vertorTermos(i) & "<BR>"
	end if
	campo = REPLACE(campo,vertorTermos(i),"")
NEXT

if resultado <> "" then
response.write "<BR><B>" & resultado & "</B><BR>"
response.End()
end if

protetorSQL = TRIM(campo)
protetorSQL = REPLACE(campo,"|"," ")


end function

' rotina de limpeza e protecao de dados enviados pra previnir sql injection
' testa formularios

	For Each Item In Request.Form
		protetorSQL(Request.Form(Item))
	Next

' testa querysting

	protetorSQL(request.servervariables("QUERY_STRING"))



%>