<%
function calculaLiquido(idVoucherTemp)
	
	set rs = objConn.execute("select * from vouchertemp where id = '"&idVoucherTemp&"'")

	if rs("pagamento") = "FA" or rs("pagamento") = "AV" or rs("pagamento") = "BO" then calculaLiquido = rs("netBRL")
	
	if rs("pagamento") = "CC" then calculaLiquido = rs("comissaoBRL") * (-1) 
	
	if rs("pagamento") = "BD" or rs("pagamento") = "FR" then calculaLiquido = 0
	
end function
%>  