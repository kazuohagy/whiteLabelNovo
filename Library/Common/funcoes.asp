<%
	Session.CodePage=65001

function trataErro(erro)
	dim retornaErro
	retornaErro = retornaErro & "Ocorreu um erro de Script:"&"<br>"
	retornaErro = retornaErro & "Erro numero="& erro.number &"<br>"
	retornaErro = retornaErro & "Descricao="& erro.description &"<br>"
	retornaErro = retornaErro & "Contexto="& erro.helpcontext &"<br>"
	retornaErro = retornaErro & "Caminho="& erro.helppath &"<br>"
	retornaErro = retornaErro & "Origem="& erro.nativeerror &"<br>"
	retornaErro = retornaErro & "Fonte="& erro.source &"<br><br>" 
	trataErro = retornaErro
end function


function data(dat,tipo,sql)
	'o parametro SQL é para indicar se a data servirá para uma string SQL...
	'caso a data seja vazia, ele retorna NULL sem aspas, senão retorna a data com aspas '0000-00-00'
	'isso é necessário para o SQL Server 2000
	dim vetdata, datafinal
	
	if isdate(dat) then	
		dat = REPLACE(dat,"/","-")
		vetdata = split(dat,"-")
	
		if tipo = 2 then ' 0000-00-00
			'sql = 0
			datafinal = vetdata(2) & "-" & right("0"&vetdata(1),2) & "-" & right("0"&vetdata(0),2)
			'datafinal = right("0"&vetdata(0),2) & "/" & right("0"&vetdata(1),2) & "/" & vetdata(2)
		elseif tipo = 1 then ' 00/00/0000
			datafinal = right("0"&vetdata(0),2) & "/" & right("0"&vetdata(1),2) & "/" & vetdata(2)
		end if
	else
		datafinal = "Null"
	end if
	
	if sql=1 and datafinal <> "Null" then
		data = "'" & datafinal & "'"
	else
		data = datafinal
	end if
end function


function forMoeda(str,decimais)
'RS para input:	replace ( replace ( formatnumber(str,2) , "." ,"") , "," ,".")
	if len(str)=0 or isnull(str) or str="" then str = 0
	decimais = cint(decimais)
	if LEFT(RIGHT(str,3),1)= "." then
	forMoeda = str
	else
	str = formatnumber(str,decimais)
	str = replace( str , "." , ""  )
	str = replace( str , "," , "." )
	forMoeda = str
	end if
end function


function achaPlano(cod)
	dim rstmp
	if isnull(cod) or cod="" or len(cod) <=0 then cod=0
	set rstmp = objConn.execute("SELECT nome FROM  planos WHERE id="&cod)
	if rstmp.eof then
		achaPlano = "<b><font color='red'>Plano " & cod & ": N/A</font></b>"
	else
		achaPlano = rstmp(0)
	end if
	set rstmp = nothing
end function

function achaCliente(cod)
	dim rstmp
	if isnull(cod) or cod="" or len(cod) <=0 then cod=0
	set rstmp = objConn.execute("SELECT fantasia FROM  cadCliente WHERE id="&cod)
	if rstmp.eof then
		achaCliente = "<b><font color='red'>Cliente " & cod & ": N/A</font></b>"
	else
		achaCliente = rstmp(0)
	end if
	set rstmp = nothing
end function


function tipoPlano(tipo,acao)

if acao = 1 then
	SELECT CASE tipo
		case "n": response.write "Nacional"
		case "i": response.write "Internacional"
		case "c": response.write "Cruzeiro"
	END SELECT
	else
	SELECT CASE tipo
		case "n": response.write "R$ "
		case "i": response.write "US$ "
		case "c": response.write "US$ "
	END SELECT
end if
end function


Function InsertUpdate(vsTexto , vUpdate ) 
    'vUpdate = 0  ->Insere
    'vUpdate = 1  ->Update
    Dim vTemp 
    Dim vTemp2 
    Dim vRetorno 
    Dim N 
    Dim vCampo() 
    Dim vValor() 

    vTemp = vsTexto
    ReDim vCampo(0)
    ReDim vValor(0)
    Do While InStr(vTemp, ";" + Chr(0)) > 0
        ReDim Preserve vCampo(UBound(vCampo) + 1)
        ReDim Preserve vValor(UBound(vValor) + 1)
        vTemp2 = Mid(vTemp, 1, InStr(vTemp, ";" + Chr(0)) - 1)
        vTemp = Mid(vTemp, InStr(vTemp, ";" + Chr(0)) + 2)
        vCampo(UBound(vCampo)) = Trim(Mid(vTemp2, 1, InStr(vTemp2, "=") - 1))
        vValor(UBound(vCampo)) = Trim(Mid(vTemp2, InStr(vTemp2, "=") + 1))
    Loop

    If Trim(vTemp) <> "" Then
        ReDim Preserve vCampo(UBound(vCampo) + 1)
        ReDim Preserve vValor(UBound(vValor) + 1)
        vCampo(UBound(vCampo)) = Mid(vTemp, 1, InStr(vTemp, "=") - 1)
        vValor(UBound(vCampo)) = Mid(vTemp, InStr(vTemp, "=") + 1)
    End If
    If vUpdate = 0 Then
        vRetorno = "( "
        For N = 1 To UBound(vCampo)
            vRetorno = vRetorno + vCampo(N) + ", "
        Next
        vRetorno = Mid(vRetorno, 1, Len(vRetorno) - 2) + " ) VALUES ( "
        For N = 1 To UBound(vCampo)
            vRetorno = vRetorno + vValor(N) + ", "
        Next
        vRetorno = Mid(vRetorno, 1, Len(vRetorno) - 2) + " ) "
    Else
        For N = 1 To UBound(vCampo)
            vRetorno = vRetorno + vCampo(N) + " = " + vValor(N) + " ,"
        Next
        vRetorno = Mid(vRetorno, 1, Len(vRetorno) - 2)
    End If

    InsertUpdate = vRetorno

End Function

'Operador ternário para ASPCLASSIC
Function TernaryOperator(condition, value_if_true, value_if_false)
    If condition Then
        TernaryOperator = value_if_true
    Else
        TernaryOperator = value_if_false
    End If
End Function

%>