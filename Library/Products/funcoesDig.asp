<%
function geraId(num)

Dim mult, numero, str, i, numfinal

	num = replace(num," ","")
	num = replace(num,".","")
	num = replace(num,"D","")
	
	mult = 2
	numero = 0
	str = ""
	
	for i = 0 to len(num)-1
		numero = numero + (mid(num,len(num)-i,1) * mult)
		
		if mult = 9 then
		mult = 2
		else
		mult = mult + 1
		end if
	next
	
	numero = 11 - (numero mod 11)
	
	if (numero = 0) or (numero = 1) or (numero > 9) then
		numero = 1
	end if
	
	numfinal = left(num,4) & numero & right(num,len(num)-4)

	geraId = numfinal

end function

function dac(num,indice)
Dim i2, numero, i, tempo, temp2, temp, numero2, i3
	num = replace(num," ","")
	num = replace(num,".","")
	num = replace(num,"D","")
	numero = 0
	
	if indice=2 or indice=3 then
		i2 = 1
	else
		i2 = 2
	end if
	
	for i = 1 to len(num)
		temp = mid(num,i,1) * i2
		if temp>9 then
			temp2 = 0
			for i3 = 1 to len(temp)
				temp2 = temp2 + cint(0&mid(temp,i3,1))
			next
			temp = temp2
		end if
		
		numero = numero + temp
		
		if i2 = 1 then
			i2 = 2
		else
			i2 = 1
		end if
	next
	
	numero2 = numero
	
	while numero2 mod 10 <> 0
		numero2 = numero2 + 1
	wend
	
	numero = numero2 - numero
	
	dac = abs(numero)
end function

function digitavelMostra(num)

digitavelMostra =MID(num,1,5) & "." & MID(num,6,5) & "  " & MID(num,11,5) & "." & MID(num,16,6) & "  " & MID(num,22,5) & "." & MID(num,27,6) & "  " & MID(num,33,1) & "  " & MID(num,34,14)

end function

function nossonumeroDig(num)
	Dim mult, numero, str, i, numfinal

	mult = 2
	numero = 0
	str = ""
	
	for i = 0 to len(num)-1
		numero = numero + (mid(num,len(num)-i,1) * mult)
		mult = mult + 1
		if mult > 7 then
		mult = 2
		end if
	next
	
	numero = 11 - (numero mod 11)
	
	if numero = 10 then 'se o numero mod 11 for 1
		nossonumeroDig =  "P"
	elseif numero = 11 then ' o numero mod 11 for zero
		nossonumeroDig =  "0"
	else 'senão...
		nossonumeroDig = cstr(numero)
	end if

end function

%>
