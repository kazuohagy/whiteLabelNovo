<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<!--#include file="../Library/Common/enviaEmail.asp" --> 
<!--#include file="../Library/Products/PriceFunctions.asp" -->
<%
modal = true


cotacaoId = request.QueryString("cotacao_id")
planoId = request.QueryString("planoId")
chave = request.QueryString("chave")

address = "../../Products/Product_Form.asp?planoId=" & planoId & "&chave=" & chave & "&"
set dados_cotacao_planoRS = objConn.execute("select cotacao_reg.*, cotacao_reg_pax.* from cotacao_reg inner join cotacao_reg_pax ON cotacao_reg_pax.cotacao_id = cotacao_reg.id INNER JOIN planos ON planos.id = cotacao_reg_pax.plano_id WHERE cotacao_reg.id = " & cotacaoId & " AND planos.id = " & planoId & " ORDER BY ordemExibicao")    

set planoRS = objConn.Execute("SELECT * FROM planos where id = " & planoId & "")
if planoRS("publicado") = 0 then acordo=1

set IdadeRS = objConn.execute("SELECT coalesce(idadeMinima50,0) as idade_Acrescimo, coalesce(idadeMinima,0) as idade_limite from planos where id = '"&planoId&"'")

    dias = dados_cotacao_planoRS("vigencia")
    categoria = dados_cotacao_planoRS("categoria_id")
    destino = dados_cotacao_planoRS("destino_id")
    nPax = dados_cotacao_planoRS("nPax_total")
    n_novo = dados_cotacao_planoRS("nPax_Novo")
    n_idoso = dados_cotacao_planoRS("nPax_idoso")
    dataInicio = dados_cotacao_planoRS("data_inicio")
    dataFim = dados_cotacao_planoRS("data_fim")
    vigencia = dados_cotacao_planoRS("vigencia")
    familiar = dados_cotacao_planoRS("familiar")
    moeda = dados_cotacao_planoRS("moeda")
    upgrade_cancel = dados_cotacao_planoRS("upgradeCancel")
    upgrade_covid = dados_cotacao_planoRS("upgradeCovid") 
	origemRecep = dados_cotacao_planoRS("origemRecep") 

    set ViagemRS = objConn.execute("SELECT * FROM viagem_destino where id = '"&destino&"' and ativo = 1;")
    temBoleto = "N"
    sessao = Session.SessionID  

    objConn.Execute("INSERT INTO emissaoProcesso (sessao,chave,usuarioLogin,usuarioIp,planoId,clienteId,dataInicio,dataFim,nPax,destino,passaAngola,cotacaoProcesso) values ( '"&sessao&"', '"&chave&"','"&request.cookies("FCNET_MIC")("login")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.QueryString("planoId")&"','"&Request.cookies("FCNET_MIC")("idAge")&"','"&dataInicio&"','"&dataFim&"','"&nPax&"','"&destino&"','"&p50&"','"&cotacaoId&"')")

	set processoRS = objConn.Execute("SELECT MAX(id) FROM emissaoProcesso WHERE chave = '"&chave&"'")
    processo = processoRS(0)        

	'historico do processo
	objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"https://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Processo iniciado na sessao "&sessao&"')")
    
	if not planoRS.eof then
		if planoRS("familiar") = 1 then familiar = 1 else familiar = 0
	else
		response.Write "<script>alert('Nenhum plano encontrado.\nFavor selecionar novamente.')</script>"
		response.End()
	end if
	
	idAge = request.cookies("FCNET_MIC")("idAge")
	
	'hitorico do processo
	objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"https://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Formulario de PAX iniciado')")
	
	Set cliRS = objConn.Execute("Select * from cadCliente where id='"&request.cookies("FCNET_MIC")("idAge")&"'")

    nPax_calc = nPax
    
    dim montaJava
    Dim agencia
    Dim cor, BG
    Dim cliRS,pax
    pax = nPax
    p = 1
    Dim i, i2, agora, p, ageNi
    Dim iniciando, nascendo
    Dim montaAnos, montagestante, fim, navegador
    navegador=left(request.QueryString("navegador"),45)

    Function trata_filer(valor,tamanho,lado,str_cpl)
        valor = TRIM(valor)
            if ISNULL(valor) then valor = " "

            valor = REPLACE(valor,CHR(9),"")
            lado = UCASE(lado)

            if isnull(valor) then valor = " "

            if isnull(str_cpl) then str_cpl = " "
            if len(str_cpl) = 0 then str_cpl = " "
        
            if lado = "E" then
                For iW = LEN(valor) to CINT(tamanho)-1
                    valor = valor & str_cpl
                NEXT
                valor = LEFT(valor,tamanho)
                                
            end if

            if lado = "D" then
                For iW = LEN(valor) to CINT(tamanho)-1
                    valor = str_cpl & valor
                NEXT		
                valor = RIGHT(valor,tamanho)
            end if
        trata_filer = valor
    end function

    function data(dat,tipo,sql)
        dim vetdata, datafinal
        
        if isdate(dat) then	
        
            if MID(dat,3,1) = "/" or MID(dat,2,1) = "/" then
                vetdata = split(dat,"/")
            else
                vetdata = split(dat,"-")
            end if

            SELECT CASE tipo
                CASE 1: datafinal = right("0"&vetdata(0),2) & "/" & right("0"&vetdata(1),2) & "/" & vetdata(2)
                CASE 2: datafinal = right("0"&vetdata(0),2) & "/" & right("0"&vetdata(1),2) & "/" & vetdata(2)
                CASE 3: datafinal = right("0"&vetdata(0),2) & "/" & right("0"&vetdata(1),2) & "/" & vetdata(2)
                CASE 4: datafinal = LEFT(vetdata(2),4) & "-" & right("0"&vetdata(1),2) & "-" & right("0"&vetdata(0),2)
                CASE 5: datafinal = LEFT(vetdata(2),4) & right("0"&vetdata(1),2) & right("0"&vetdata(0),2)
            end select
        else
            datafinal = "Null"
            
        end if
        
        if sql=1 and datafinal <> "Null" then
            data = "'" & datafinal & "'"
        else
            data = datafinal
        end if
        
    end function

    function montarGestacao(gesZ)
        montagestante = "<OPTION value='0' disabled selected>Selec.</OPTION>"
        iniciando = 1
        nascendo = 46
        i=iniciando
        while i < nascendo
            montagestante = montagestante & "<option value='" & i & "'"
            if gesZ <> "" then 
                        if cint(gesZ) = cint(i) then montagestante = montagestante & " selected "
            end if
            montagestante = montagestante & ">" & i & " semana(s)" 
            montagestante = montagestante & "</option>" & Chr(13)
            i=i+1
        wend
        montarGestacao = montagestante    
    end function

    function montarAnosAgora(anoX)
        montaAnos = "<OPTION disabled selected value=''>Ano</OPTION>"
        agora = Year(now())
        fim = agora + 6
        i=agora
        while i>(1900 + (54 + (Year(now()) - 2020)))
            montaAnos = montaAnos & "<option value='" & i & "'"
            if anoX <> "" then 
                        if cint(anoX) = cint(i) then montaAnos = montaAnos & " selected "
            end if
            montaAnos = montaAnos & ">" & i
            montaAnos = montaAnos & "</option>" & Chr(13)
            i=i-1
        wend
        montarAnosAgora = montaAnos    
    end function

    function montarAnosIdoso(anox)
        montaAnosIdoso = "<OPTION disabled selected value='' >Ano</OPTION>"
        agora = Year(now()) - 65
        fim = agora + 6
        i=agora
        while i>(1900 + (33 + (Year(now()) - 2020)))
            montaAnosIdoso = montaAnosIdoso & "<option value='" & i & "'"
            if anoY <> "" then 
                        if cint(anoX) = cint(i) then montaAnosIdoso = montaAnosIdoso & " selected "
            end if
            montaAnosIdoso = montaAnosIdoso & ">" & i
            montaAnosIdoso = montaAnosIdoso & "</option>" & Chr(13)
            i=i-1
        wend
        montarAnosIdoso = montaAnosIdoso    
    end function
    dim mesInicio, mesFim, diaInicio, diaFim, anoInicio, anoFim
    dataInicio = REPLACE(data(CDATE(dados_cotacao_planoRS("data_inicio")),5,0),"-","")
        data_1 = LEFT (dataInicio,4)
        data_2 = MID (dataInicio,5,2)
        data_3 = RIGHT (dataInicio,2)	
        'Recuperando as datas'
        diaInicio = data_3
        mesInicio = data_2
        anoInicio = data_1
    dataInicio = data_3 & "-" & data_2 & "-" & data_1

    dataFim = REPLACE(data(CDATE(dados_cotacao_planoRS("data_fim")),5,0),"-","")
        data_1 = LEFT (dataFim,4)
        data_2 = MID (dataFim,5,2)
        data_3 = RIGHT (dataFim,2)	
        'Recuperando as datas'		
        mesFim = data_2
        diaFim = data_3
        anoFim = data_1
    dataFim = data_3 & "-" & data_2 & "-" & data_1

    'Converte mes em numero para nome do mes em ASP CLASSIC
    Function ConvertMes(mes)
        Select Case mes
            Case "01": ConvertMes = "Janeiro"
            Case "02": ConvertMes = "Fevereiro"
            Case "03": ConvertMes = "Março"
            Case "04": ConvertMes = "Abril"
            Case "05": ConvertMes = "Maio"
            Case "06": ConvertMes = "Junho"
            Case "07": ConvertMes = "Julho"
            Case "08": ConvertMes = "Agosto"
            Case "09": ConvertMes = "Setembro"
            Case "10": ConvertMes = "Outubro"
            Case "11": ConvertMes = "Novembro"
            Case "12": ConvertMes = "Dezembro"
        End Select
    End Function
    'Fim da função

    if request.cookies("FCNET_MIC")("logado") = 1 then
        objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"https://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Verificacao de login. Usuário "&request.cookies("FCNET_MIC")("login")&" logado')")
    else
        objconn.execute("INSERT INTO emissaoProcessoHistorico (processoId,url,obsTxt) VALUES ('"&processo&"','"&"https://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")&"','Verificacao de login. Usuário não logado')")
    end if    
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Complete os detalhes de sua compra</title>
    <!--#include file="../Components/HTML_Head2.asp" -->
    <script src="../JavaScript/jquery.mask.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>    
    <script LANGUAGE="LiveScript">

        function check_cpf_SUSEP(numcpf,cmp,numId){
            if ( (numcpf=="11111111111")  || (numcpf=="22222222222")  || (numcpf=="33333333333") || (numcpf=="44444444444") || (numcpf=="55555555555") || (numcpf=="66666666666") || (numcpf=="77777777777") || (numcpf=="88888888888") || (numcpf=="99999999999") || (numcpf=="00000000000") || numcpf.length < 11)  {
        
                alert("CPF invalido!");
                document.getElementById(numId).value="";
     
            }else{
                x = 0;
	            soma = 0;
	            dig1 = 0;
	            dig2 = 0;
	            texto = "";
	            numcpf1="";
	            len = numcpf.length; x = len -1;
	
	            // var numcpf = "12345678909";
	            for (var i=0; i <= len - 3; i++) {
		            y = numcpf.substring(i,i+1);
		            soma = soma + ( y * x);
		            x = x - 1;
		            texto = texto + y;
                }
            
	            dig1 = 11 - (soma % 11);
	
		        if (dig1 == 10) dig1=0 ;
		        if (dig1 == 11) dig1=0 ;
		
	            numcpf1 = numcpf.substring(0,len - 2) + dig1 ;
	            x = 11; soma=0;
	
	            for (var i=0; i <= len - 2; i++) {
		            soma = soma + (numcpf1.substring(i,i+1) * x);
		            x = x - 1;
	            }
	
	            dig2= 11 - (soma % 11);
	
	            if (dig2 == 10) dig2=0;
	            if (dig2 == 11) dig2=0;
	
	            //alert ("Digito Verificador : " + dig1 + "" + dig2);
	            if ((dig1 + "" + dig2) == numcpf.substring(len,len-2)) {
		            return true;
	            }
	
	            alert ("Numero do CPF invalido !!!");
	            document.getElementById(numId).value = "";
                return false;
            }
        }

        function limpa_formulario_cep() {
        //Limpa valores do formulario de cep.
            document.getElementById('enderecoPax').value=("");
            document.getElementById('bairroPax').value=("");
            document.getElementById('cidadePax').value=("");
            document.getElementById('ufPax').value=("");
        }
	
        function meu_callback(conteudo) {
            if (!("erro" in conteudo)) {
			//Atualiza os campos com os valores.
                if (conteudo.logradouro == "" || conteudo.bairro == "" || conteudo.localidade == "") {
                    document.getElementById('enderecoPax').readOnly= false;
                    document.getElementById('bairroPax').readOnly= false;
                    document.getElementById('cidadePax').readOnly= false;
                }
                document.getElementById('enderecoPax').value=(conteudo.logradouro);
                document.getElementById('bairroPax').value=(conteudo.bairro);
                document.getElementById('cidadePax').value=(conteudo.localidade);
                document.getElementById('ufPax').value=(conteudo.uf);
            }   //end if.
            else {
            //CEP nao Encontrado.
                limpa_formulario_cep();
                alert("CEP nao encontrado.");
                document.getElementById('enderecoPax').readOnly= false;
                document.getElementById('bairroPax').readOnly= false;
                document.getElementById('cidadePax').readOnly= false;
                document.getElementById('ufPax').readOnly= false;
            }
        }
        
        function pesquisacep(valor) {
        //Nova variavel "cep" somente com digitos.
            var cep = valor.replace(/\D/g, '');

            //Verifica se campo cep possui valor informado.
            if (cep != "") {
            //Expressao regular para validar o CEP.
                var validacep = /^[0-9]{8}$/;

                //Valida o formato do CEP.
                if(validacep.test(cep)) {

                //Preenche os campos com "..." enquanto consulta webservice.
                    document.getElementById('enderecoPax').value="";
                    document.getElementById('bairroPax').value="";
                    document.getElementById('cidadePax').value="";
                    document.getElementById('ufPax').value="";

                    //Cria um elemento javascript.
                    var script = document.createElement('script');

                    //Sincroniza com o callback.
                    script.src = '//viacep.com.br/ws/'+ cep + '/json/?callback=meu_callback';

                    //Insere script no documento e carrega o conteudo.
                    document.body.appendChild(script);

                    //Trava os elementos, e deixa como readonly
                    document.getElementById('enderecoPax').readOnly= true;
                    document.getElementById('bairroPax').readOnly= true;
                    document.getElementById('cidadePax').readOnly= true;
                    document.getElementById('ufPax').readOnly= true;

                } //end if.
                else {
                //cep e invalido.
                    limpa_formulario_cep();
                    alert("Formato de CEP invalido.");
                    document.getElementById('enderecoPax').readOnly= false;
                    document.getElementById('bairroPax').readOnly= false;
                    document.getElementById('cidadePax').readOnly= false;
                    document.getElementById('ufPax').readOnly= false;
                }
            } //end if.
            else {
            //cep sem valor, limpa formulario.
                limpa_formulario_cep();
            }
        };
    </script>
    <SCRIPT SRC="../JavaScript/fcCalculos.js"></script>
    <SCRIPT SRC="../JavaScript/fcJanelas.js"></script>
    <SCRIPT SRC="../JavaScript/fcGeral.js"></script>
    <script language="JavaScript">

            <%
                For p = 1 to pax
            %>
            $(document).ready(function(){
                
                $('#DDDcelularPax<%=p%>').mask('(00) 00000-0000');
                $('#cepPax').mask('00000-000');
                $('.telefoneP ').mask('(00)000000000');
                $('#FoneN').mask('(00)000000000');  
            });

            <%
                next
            %>

            $(document).ready(function(){
   
                $('.FoneN').mask('(00)000000000');
                $('#FoneN').mask('(00)000000000');  
            });

            function fixaNum(objeto) { //validate3
	            var keypress = event.keyCode; 
	            var campo = eval (objeto);
	            var sCaracteres = '0123456789';

	            if (sCaracteres.indexOf(String.fromCharCode(keypress))!=-1){
		            event.returnValue = true;
	            }else
	                event.returnValue = false;
            }

            function confere_reserva(){
	            	
	
	            <%
		            For p=1 to pax
	            %>

	                var docPax<%=p%> = document.form1.docPax<%=p%>.value
	
	                if (docPax<%=p%>=="" || docPax<%=p%>=="00000000000"|| docPax<%=p%>=="11111111111" || docPax<%=p%>=="22222222222" || docPax<%=p%>=="33333333333" || docPax<%=p%>=="44444444444" || docPax<%=p%>=="55555555555" || docPax<%=p%>=="66666666666" || docPax<%=p%>=="77777777777" || docPax<%=p%>=="88888888888" || docPax<%=p%>=="99999999999" || docPax<%=p%>.length < 11) {
		                alert("\Preencha corretamente o campo DOCUMENTO do passageiro <%=p%>.");
		                document.form1.docPax<%=p%>.style.backgroundColor = "#FFCCCC";
		                document.form1.docPax<%=p%>.focus()
		                return false
	                }	
	
	                if ((document.form1.day<%=p%>.value!= "xx") || (document.form1.month<%=p%>.value!= "xx") || (document.form1.year<%=p%>.value!= "xx")){
		                var idadePax<%=p%> = document.form1.idadePax<%=p%>.value
			            if ((idadePax<%=p%>=="") || (idadePax<%=p%><0) || (idadePax<%=p%>=="NaN")){
				            alert("\Insira a data de nascimento do pax <%=p%>.");
				            document.form1.day<%=p%>.style.backgroundColor = "#FFCCCC";
				            document.form1.day<%=p%>.focus()
				        return false
			            }
			
		            var day<%=p%> = document.form1.day<%=p%>.options[document.form1.day<%=p%>.selectedIndex].value;
		            var month<%=p%> = document.form1.month<%=p%>.options[document.form1.month<%=p%>.selectedIndex].value;
		            var year<%=p%> = document.form1.year<%=p%>.options[document.form1.year<%=p%>.selectedIndex].value;
		
		            if ((day<%=p%>=="xx") || (month<%=p%>=="xx") || (year<%=p%>=="xx")){
			            alert("\Insira corretamente a data de nascimento do pax <%=p%>.");
			            document.form1.day<%=p%>.focus()
			            return false
		            }
	            }
	
	            <%
		            next
	            %>
	
            }

            function verificaGestante(idadePax,sexoPax){
	            if (( parseFloat(idadePax) > 40 ) || (sexoPax != 'F') ){	
		            alert('\nPara cobertura adicional de gravidez:\n\nApenas pax do sexo feminino com idade maxima de 40 anos')
		            return false;	
	            }
	            else{
		            return true;
	            }
            }

            
function confere_emissao(){

function roda(){
    document.form1.aguarde.value=document.form1.aguarde.value+'.'
    x=x+1
    
    if (x<10){
        setTimeout(roda,300)
        
    }
    else{
        x=0
        document.form1.aguarde.value='Aguarde'
        setTimeout(roda,300)
        
    }
}
<%
    For p=1 to pax
%>
if (document.getElementById('gestantePax<%=p%>').value == "Sim"){
    if (document.getElementById('semanasGes<%=p%>').value == ""){
        alert("\nInforme as semanas de gestacao.");
        document.form1.semanasGes<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.semanasGes<%=p%>.focus();
        return false
    }
}

if (document.getElementById('gestantePax<%=p%>').value == "Sim"){
    if (document.getElementById('semanasGes<%=p%>').value >= 34){
        alert("\nLimite de semanas para emissao � de 33 semenas de gesta��o");
        document.form1.semanasGes<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.semanasGes<%=p%>.focus();
        return false
    }
}

if (document.getElementById('gestantePax<%=p%>').value == "Sim"){
    if (document.getElementById('idadePax<%=p%>').value > 43){
        alert("\nLimite de idade para emiss�o de gestantes � de 43 anos");
        document.form1.idadePax<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.idadePax<%=p%>.focus();
        return false
    }
}

var docPax<%=p%> = document.form1.docPax<%=p%>.value
if($('#estrangeiro<%=p%>').val() != 'EST'){
    if (docPax<%=p%>=="" || docPax<%=p%>=="00000000000"|| docPax<%=p%>=="11111111111" || docPax<%=p%>=="22222222222" || docPax<%=p%>=="33333333333" || docPax<%=p%>=="44444444444" || docPax<%=p%>=="55555555555" || docPax<%=p%>=="66666666666" || docPax<%=p%>=="77777777777" || docPax<%=p%>=="88888888888" || docPax<%=p%>=="99999999999" || docPax<%=p%>.length < 11) {
        alert("\Preencha corretamente o campo DOCUMENTO do passageiro <%=p%>.");
        document.form1.docPax<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.docPax<%=p%>.focus()
        return false

    }	
}
var nomePax<%=p%> = document.form1.nomePax<%=p%>.value
    if ((nomePax<%=p%>=="") || (nomePax<%=p%>.length < 2 ) || (nomePax<%=p%>.indexOf('  ', 0) != -1) || (nomePax<%=p%>.charAt(0) ==" ")){
        alert("\Preencha corretamente o campo NOME do passageiro <%=p%>.");
        document.form1.nomePax<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.nomePax<%=p%>.focus();
        return false
    }
var sobrePax<%=p%> = document.form1.sobrePax<%=p%>.value
    if ((sobrePax<%=p%>=="") || (sobrePax<%=p%>.length < 2 ) || (sobrePax<%=p%>.indexOf('  ', 0) != -1) || (sobrePax<%=p%>.charAt(0) ==" ")){
        alert("\Preencha corretamente o campo SOBRENOME do passageiro <%=p%>.");
        document.form1.sobrePax<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.sobrePax<%=p%>.focus();
        return false
    
    }

var sexoPax<%=p%> = document.form1.sexoPax<%=p%>.options[document.form1.sexoPax<%=p%>.selectedIndex].value;
    if (sexoPax<%=p%>==""){
        alert("\nSelecione o SEXO do passageiro <%=p%>.");
        document.form1.sexoPax<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.sexoPax<%=p%>.focus()
        return false

    }

var day<%=p%> = document.form1.day<%=p%>.value
    if (day<%=p%>=="0" || day<%=p%>=="" ){
        alert("\Preencha corretamente o campo DIA do passageiro <%=p%>.");
        document.form1.day<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.day<%=p%>.focus();
        return false
    
    }

var month<%=p%> = document.form1.month<%=p%>.value
    if (month<%=p%>=="0" || month<%=p%>=="" ){
        alert("\Preencha corretamente o campo MES do passageiro <%=p%>.");
        document.form1.month<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.month<%=p%>.focus();
        return false
    }

var year<%=p%> = document.form1.year<%=p%>.value
    if (year<%=p%>=="0" || year<%=p%>=="" ){
        alert("\Preencha corretamente o campo ANO do passageiro <%=p%>.");
        document.form1.year<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.year<%=p%>.focus();
        return false
    }
var idadePax<%=p%> = document.form1.idadePax<%=p%>.value
    if (idadePax<%=p%>=="" || idadePax<%=p%> == "NaN"){
        alert("\Preencha corretamente o campo ANO DO NASCIMENTO do passageiro <%=p%>.");
        document.form1.year<%=p%>.style.backgroundColor = "#FFCCCC";
        document.form1.year<%=p%>.focus();
        return false
    
    }

<%
    next
%>

var cepPax = document.form1.cepPax.value
    if (cepPax==""){
        alert("\Preencha o campo CEP do passageiro.");
        document.form1.cepPax.style.backgroundColor = "#FFCCCC";
        document.form1.cepPax.focus();
        return false;
    }

var enderecoPax = document.form1.enderecoPax.value
    if (enderecoPax==""){
        alert("\Preencha o campo ENDERECO do passageiro.");
        document.form1.enderecoPax.style.backgroundColor = "#FFCCCC";
        document.form1.enderecoPax.focus();
        return false;
    }

var numeroPax = document.form1.numeroPax.value
    if (numeroPax==""){
        alert("\Preencha o campo NUMERO do endereco do passageiro.");
        document.form1.numeroPax.style.backgroundColor = "#FFCCCC";
        document.form1.numeroPax.focus();
        return false;
    }

var bairroPax = document.form1.bairroPax.value
    if (bairroPax==""){
        alert("\Preencha o campo BAIRRO do endereco do passageiro.");
        document.form1.bairroPax.style.backgroundColor = "#FFCCCC";
        document.form1.bairroPax.focus();
        return false;
    }

var cidadePax = document.form1.cidadePax.value
    if (cidadePax==""){
        alert("\Preencha o campo CIDADE do passageiro.");
        document.form1.cidadePax.style.backgroundColor = "#FFCCCC";
        document.form1.cidadePax.focus()
        return false
    }

var ufPax = document.form1.ufPax.options[document.form1.ufPax.selectedIndex].value;
    if (ufPax==""){
        alert("\nSelecione o Estado.");
        document.form1.ufPax.style.backgroundColor = "#FFCCCC";
        document.form1.ufPax.focus()
        return false
    }

var foneN = document.form1.foneN.value
    if (foneN==""){
        alert("\Preencha o campo FONE do passageiro.");
        document.form1.foneN.style.backgroundColor = "#FFCCCC";
        document.form1.foneN.focus()
        return false
    }

if (document.form1.contatoNome.value=="" && 1==2){
    alert("\Preencha o campo nome do contato.");
    document.form1.contatoNome.focus()
    return false
}

if (document.form1.contatoEndereco.value=="" && 1==2){
    alert("\Preencha o campo endereço do contato.");
    document.form1.contatoEndereco.focus()
    return false
}

if (document.form1.contatoFoneN.value=="" && 1==2){
    alert("\Preencha o campo telefone do contato.");
    document.form1.contatoFoneN.focus()
    return false
}

var x=0
var numeroPax_emite = <%=nPax%> 

//document.form1.submit2.style.display='none'
document.form1.aguarde.style.display=''
document.getElementById('salvarSair').style.display='none';
var x=0
roda();

} 
////fim da funcao confere_emissao()       
    function removeAspa(campo){
	    campo.value=campo.value.replace("'","").replace("'","").replace("'","").replace("'","").replace("'","").replace("'","").replace("'","")
    }

    function formatanumero(numero,decimais){ 
	     var num = parseFloat(numero);
	    var casas = Math.pow(10, decimais);
	    var result = Math.round( num*casas)/casas;
	    var result = result.toString();

	    if(result.indexOf(".") != -1){
		    if(result.split('.')[1].length == '1'){
                result = result + '0'
            }
		        if(result=='NaN'){
                    result='0.00'
                }
	        }
	        else{
		        return result + '.00';
	        }

	        return result; 
        }
    </script>
    <script type="text/javascript">

        function verIDade(pax){
            document.getElementById('alerta_idade_'+pax).style.display='none';
            var idadePax = document.getElementById("idadePax" + pax).value	

            if ((parseFloat(idadePax) > <%=IdadeRS("idade_limite")%>) && (<%=IdadeRS("idade_limite")%> != 0))
            {			
                document.getElementById("idadePax" + pax).value = '';
                document.getElementById("year" + pax).value = '';
                document.getElementById('alerta_idade_'+pax).style.display='';
                document.getElementById('alerta_idade_'+pax).innerHTML = "<br>Idade máxima para emissão <%=planoRS("idadeMinima")%> anos";
            }
        }
        
        function verUpgrade(pax){
            if (document.getElementById("planoCancelamento"+pax)){
                var cancelValue = parseFloat(document.getElementById("planoCancelamento"+pax).value);
                if($("#planoCancelamento"+pax+" option:selected").attr("id"))
                $("#planoCancelamento_id"+pax).val($("#planoCancelamento"+pax+" option:selected").attr("id"));
            }            
            else{
                var cancelValue = 0;
            }
            if (document.getElementById("planoCovid"+pax)){
                var covidValue = parseFloat(document.getElementById("planoCovid"+pax).value);
                if($("#planoCovid"+pax+" option:selected").attr("id"))
                $("#planoCovid_id"+pax).val($("#planoCovid"+pax+" option:selected").attr("id"));
            }
            else{
                var covidValue = 0;
            }

            var cambio = parseFloat(document.getElementById("cambio" + pax).value);
            var valorUSD = parseFloat(document.getElementById("valorUSD_Origin" + pax).value);
            var valorBRL = parseFloat(document.getElementById("valorBRL_Origin" + pax).value);
            
            document.getElementById("valorUSD" + pax).value  = formatanumero(valorUSD + cancelValue + covidValue,2);
            <% if planoRS("nacionalidade") = "i" then %>                
                document.getElementById("valorBRL" + pax).value  = formatanumero(valorBRL + ((cancelValue + covidValue) * cambio),2);
            <% end if %>

            compoeValor();
        }

        $(document).ready(function () {
            var npax = <%=nPax%>;

            for (i = 1; i <= npax; i++) {
                verUpgrade(i);                
            }
        });
        function formatarReais(valor) {
            return new Intl.NumberFormat('br-BR', { style: 'currency', currency: 'BRL' }).format(valor);
        }

        function compoeValor(){
            //atualiza valor do processo
            var somaTodos = 0
            for (t=1;t<=<%=nPax%>;t++){
                somaTodos = somaTodos + parseFloat(document.getElementById("<% if planoRS("nacionalidade") = "i" then response.write "valorBRL" else response.write "valorUSD" end if %>"+t).value );
            }
            document.getElementById('valorParParcelar').value = formatanumero(somaTodos,2);
            $("#luigi").text(formatarReais(formatanumero(somaTodos,2)));
            $("#valorParcelado").text(formatarReais(formatanumero(somaTodos/10,2)));
            $("#parcelas").text("10x");
            
            $("#porPessoa").text(formatarReais(formatanumero(somaTodos/<%=nPax%>,2)));
            // if(formatanumero(somaTodos,2) < 60) {
            //     $("#valorParcelado").text(formatarReais(formatanumero(somaTodos,2)));
            //     $("#parcelas").text("1x");
            // }
            // else if(formatanumero(somaTodos,2) >= 60 && formatanumero(somaTodos,2) < 90){
            //     $("#valorParcelado").text(formatarReais(formatanumero(somaTodos/2,2)));
            //     $("#parcelas").text("2x");
            // }
            // else if(formatanumero(somaTodos,2) >= 90 && formatanumero(somaTodos,2) < 120){
            //     $("#valorParcelado").text(formatarReais(formatanumero(somaTodos/3,2)));
            //     $("#parcelas").text("3x");
            // }
            // else if(formatanumero(somaTodos,2) >= 120 && formatanumero(somaTodos,2) < 150){
            //     $("#valorParcelado").text(formatarReais(formatanumero(somaTodos/4,2)));
            //     $("#parcelas").text("4x");
            // }
            // else if(formatanumero(somaTodos,2) >= 150 && formatanumero(somaTodos,2) < 1000){
            //     console.log("raluca");
            //     $("#valorParcelado").text(formatarReais(formatanumero(somaTodos/5,2)));
            //     $("#parcelas").text("5x");
            // }
            // else{
            //     $("#valorParcelado").text(formatarReais(formatanumero(somaTodos/10,2)));
            //     $("#parcelas").text("10x");
            // }
                    
        }

        function func_fem(pax){	
            valor = document.getElementById("sexoPax"+pax).value;
            if (valor == 'F'){
                document.getElementById("gestacao"+pax).style.display = '';
            }
            else{

                document.getElementById('gestantePax'+pax).value = 'Nao';

                document.getElementById('semanasGes'+pax).value = '0';
    
                document.getElementById("gestacao"+pax).style.display = 'none';
                
                document.getElementById("dados_gestacao"+pax).style.display = 'none';
                
            }		
        }

        function func_gravidez(pax){

        valor = document.getElementById("gestantePax"+pax).value;
            if (valor == 'Sim'){	
                document.getElementById("dados_gestacao"+pax).style.display = '';
            }
            else{
                document.getElementById("dados_gestacao"+pax).style.display = 'none';
            }			
    }
         //os métodos abaixo sao referentes ao cupom de desconto
        function validaCupom(cp) {

            if (cp == "") {
                $("#cupom").addClass("is-invalid");
                alert("Preencha um cupom válido");
                event.preventDefault();
            }

	        var cpInput = $("#cupom");
            var processoId = $("#processo").val();
            
		    var dados = {
			    cupom: cp,
                plano: <%=request.querystring("planoId")%>,
			    processo: processoId,
                brl: $("#valorBRL1").val(),
                usd: $("#valorUSD1").val(),
                totalBRL: $("#valorParParcelar").val()
            }
            
		    $.ajax({
                url: "valida_cupom.asp",
                type: "post",
                data: dados ,
                success: function (response) {
                    
                    if (response.erro == null) { 
                        if (response.antigoBRL <= 0)
                        {
                            var antigoBRL = parseFloat(response.antigoUSD.replace(",", ".")).toFixed(2);
                
                            var novoBRLRight = parseFloat(response.novoUSD.replace(",", ".")).toFixed(2);
                        }    
                        else 
                        {
                            var antigoBRL = parseFloat(response.antigoBRL.replace(",", ".")).toFixed(2);
                
                            var novoBRLRight = parseFloat(response.novoBRL.replace(",", ".")).toFixed(2);
                        }

                        var novoTotalBRL = parseFloat(response.totalBRLdesconto.replace(",", ".")).toFixed(2);
                        
                        var novoPercentComissao = (parseInt(response.comissaoPercent) - parseInt(response.percentual));
                        
                        var totalComissao = parseFloat(response.totalComissao.replace(",", ".")).toFixed(2);

                        var novoTotalComissao = parseFloat(response.novoTotalComissao.replace(",", ".")).toFixed(2);

                        $("#cupomZone").css("display", "none");

                        $("#afterDiscount").html(`
                        
                            <div style="text-align:center;">
                                <BR>
                                <h5 style="color:green">O cupom ${response.cupom} de ${response.percentual} % foi aplicado a compra</h5>
                                <BR>

                                <table class="table table-hover">
                                    <tr class="thead-dark">
                                        <th>Valor por pax (sem acréscimo)</th>
                                        <th>Valor total</th>
                                        <th></th>
                                    </tr>
                                    <tr style="color:red;text-decoration:line-through;">
                                        <td>R$ ${antigoBRL}</td>
                                        <td>R$ ${dados.totalBRL}</td>
                                        <td></td>
                                    </tr>
                                    <tr style="color:green;">
                                        <td>R$ ${novoBRLRight}</td>
                                        <td>R$ ${novoTotalBRL}</td>
                                        <td></td>
                                    </tr>
                                </table>
                                <BR>
                                <input type="checkbox" name="acordoCupom" required="required">
                                <label> Concordo com os descontos apresentados acima</label>
                            </div>

                            
                        `);

                        $(".afterCupom").each(function()    {

                            $(this).html(" -" + response.percentual + "%");

                        });
                    }
                    else
                    {
                        $("#cupom").addClass("is-invalid");
                        alert(" O cupom " + response.erro + " está inativo, expirado ou incorreto.");
                    }
                },
                error: function(jqXHR, textStatus, errorThrown) {

			        $("#cupom").addClass("is-invalid");
        	        alert("cupom inválido");

                }
            });

        }

        // fim cupom de desconto
    </script>
    <script>
        var http = false;
            if(navigator.appName == "Microsoft Internet Explorer"){
                http = new ActiveXObject("Microsoft.XMLHTTP");
            }else{
                http = new XMLHttpRequest();
            }
            function checkImageLink(logoUrl) {
      //alterar logo  caso não exista
      var img = new Image();
      img.onload = function() {
        document.getElementById("logo-marca").src = logoUrl;
        
      };

      img.onerror = function() {
        document.getElementById("logo-marca").src = "./img/logo-parceiro.jpg";
      };

      img.src = logoUrl;
    }

    // Chame a função passando o link da imagem que você deseja verificar
    document.addEventListener("DOMContentLoaded", function() {
      var logoUrl = "https://www.nextseguro.com.br/NextSeguroViagem/old/img/agtaut/<%=request.Cookies("wlabel")("revId")%>.jpg";
     
      checkImageLink(logoUrl);
    });

    // Inicio verificação de banner
    // var bannerUrl = "https://seguroviagemnext.com.br/v2/Products/uploads/<%=request.Cookies("wlabel")("revId")%>.jpg";
    // var image = new Image();

    // image.onload = function() {
    //   document.getElementById("imagemCapa").style.backgroundImage = "url('" + bannerUrl + "')";
    //   console.log("O caminho da imagem de fundo é válido.");
    // };

    // image.onerror = function() {
    //   document.getElementById("imagemCapa").style.backgroundImage = "url('https://seguroviagemnext.com.br/v2/img/banner.png')";
    //   console.log("O caminho da imagem de fundo não é válido.");
    // };

    // image.src = bannerUrl;
    //fim da verificação de banner
    </script>
    <style>
        .checkboxed{
            left: 10px;
            top: 5px;
        }
        .grantorino{
            height: 300px !important;
            padding: 50px 40px !important;
        }
        .formHeader1{
            margin: -25px;
            padding:9px 15px;
            font-size: 20px;
            color:white !important;
            border-bottom:1px solid #eee;
            background-color: #ad0000;
            -webkit-border-top-left-radius: 5px;
            -webkit-border-top-right-radius: 5px;
            -moz-border-radius-topleft: 5px;
            -moz-border-radius-topright: 5px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
        }
        .formHeader2{
            margin: -25px;
            padding:9px 15px;
            font-size: 20px;
            color: white;
            border-bottom:1px solid #eee;
            background-color: #292d3c;
            -webkit-border-top-left-radius: 5px;
            -webkit-border-top-right-radius: 5px;
            -moz-border-radius-topleft: 5px;
            -moz-border-radius-topright: 5px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
        }

        .btnAdapt {
            font-family: 'geoMedium';
            border: 2px solid  #ad0000;
            background-color: white;
            color: black;
            font-size: 16px;
            cursor: pointer;
        }

        .tev {
            font-family: 'geoMedium';
            border-color:  #ad0000;
            color:  #ad0000;
        }

        .tev:hover {
            font-family: 'geoMedium';
            background-color:   #ad0000;
            color: white;
        }
        /* Estilo personalizado para o label "Dt de Nasc." 
        .input-group-text {
        background-color: #f0f0f0;
        color: #333;
        font-weight: bold;
        padding: 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
        }*/
        .accordion-button:not(.collapsed) {
            background-color:  #fff;
            border-color: #fff;
        }
        .input-group2 {
    position: relative;
    display: flex;
    flex-wrap: wrap;
    align-items: stretch;

}


    </style>
</head>

<body id="affinity-page">
    <header>
        <!--#include file ="../Components/Header.asp"-->
    </header>
    <!-- Início banner -->
     <!--#include file ="../Components/Banner.asp"-->
    <!-- Fim banner -->

    <div class="container">

    <div class="text-center mt-5">
        <h2>Passo 2 - Informar Dados</h2>
    </div>
    <ul class="progressbar d-flex justify-content-center mb-5">
        <li class="active">Seleção do Plano</li>
        <li class="active">Informar Dados</li>
        <li>Realizar Pagamento</li>
    </ul>
    <section class="">
        <form action="formaPagamento.asp" method="post" name="form1" id="form1" onSubmit="document.form1.emitir.disabled=true"  autocomplete="off">
            
        <br>
        <%        
            total = 0
            totalBR = 0
            total_original = 0

            WHILE NOT dados_cotacao_planoRS.EOF
                if n_novo <> 0 then
                    tarifapax = dados_cotacao_planoRS("tarifa_USD")
                    tarifapaxBR = dados_cotacao_planoRS("tarifa_BRL")
                    tarifapax_original = dados_cotacao_planoRS("tarifa_original")                    

                    For i=0 To n_novo-1
                        tarifapax_temp = dados_cotacao_planoRS("tarifa_USD")
                        tarifapaxBR_temp = dados_cotacao_planoRS("tarifa_BRL")
                        
                        total = total + tarifapax_temp
                        totalBR = total + tarifapaxBR_temp
                        total_original = total_original + tarifapax_original

                        dados_cotacao_planoRS.MOVENEXT                                                                    
                    Next                                     
                end if

                if n_idoso <> 0 then
                    tarifapax_idoso = dados_cotacao_planoRS("tarifa_USD")
                    tarifapaxBR_idoso = dados_cotacao_planoRS("tarifa_BRL")
                    tarifapax_original_idoso = dados_cotacao_planoRS("tarifa_original")

                    For i=0 To n_idoso-1
                        tarifapax_temp = dados_cotacao_planoRS("tarifa_USD")
                        tarifapaxBR_temp = dados_cotacao_planoRS("tarifa_BRL")

                        total = total + tarifapax_temp
                        totalBR = total + tarifapaxBR_temp
                        total_original = total_original + tarifapax_original_idoso

                        dados_cotacao_planoRS.MOVENEXT
                    Next
                end if
            wend

            For p=1 to nPax
        %>
        <!-- Inicio do form novo -->
        <div class="">
            <div class="card plano mb-3">
                <div class="card-body py-4 px-4">
                    <h5 class="card-title"><i class="fas fa-user"></i> 
                    <%if p <= n_novo then%>
                
                        <strong>Passageiro <%=p%></strong> 
            
                    <%else%>
                
                        <strong>Passageiro Senior <%=p-n_novo%></strong> 
                
                    <%end if%>  
                    </h5>
                    <br>
                    
                    <div class="input-group col-md-3 mb-3">
                        <div class="input-group-prepend d-flex">
                            <div class="input-group-text">
                                <label >Nacionalidade:</label>
                            </div>
                           
                            <div class="form-check checkboxed" style="margin-left:10px; display: flex;align-items: center; justify-content: center;">
                                <input type="radio" id="brasileiro<%=p%>" name="nac_pax<%=p%>" class="form-check-input" value="BRA" checked  onClick="changeNation('CPF',<%=p%>)" style="margin-right:5px" >
                                <label class="form-check-label" for="brasileiro<%=p%>" >Brasileiro </label>
                            </div>
                            <%if request.cookies("FCNET_MIC")("login") = "belv" or Request.QueryString("planoId") = 11760 or Request.QueryString("planoId") = 11761 or Request.QueryString("planoId") = 11762 or categoria = 7 then%>
                            <div class="form-check checkboxed" style="margin-left: 10px; display: flex;align-items: center; justify-content: center;">
                                <input type="radio" id="estrangeiro<%=p%>" name="nac_pax<%=p%>" class="form-check-input" value="EST" <%if nac_pax="EST"  then response.write " checked " %>  onClick="changeNation('PASS',<%=p%>)" style="margin-right:5px">
                                <label class="form-check-label" for="estrangeiro<%=p%>"> Estrangeiro </label>
                            </div>
                            <%end if%>
                           
                            <div class="input-group col-md-4 " style="justify-content:center; "hidden="hidden" disabled="disabled" name ="planoRecep<%=p%>" id="planoRecep<%=p%>">
                           
                                    <label class="input-group-text" for="inputGroupSelect05">País Emissor do Pass.</label>
                              
                                <select class=" campo-customizado" id="origemPass<%=p%>" name="origemPass<%=p%>" hidden="hidden" disabled="disabled">
                                    <option disabled selected value="">Selecione um pais</option>
                                    <%
                                        set rsPais = objConn.execute("select * from paisesLista order by id")
                                        while not rsPais.eof
                                            Response.Write "<option value="&rsPais("cod")&">"&rsPais("nome")&"</option>"
                                            rsPais.movenext
                                        wend
                                    %>
                                </select>
                            </div> 
                        </div>
                    </div>
                    <div class="col-md-2"></div>

                   
                    

                    <div class="row">
                        <div class="col-md-3 mb-3">
                            <input name="nomePax<%=p%>" type="text" id="nomePax<%=p%>" aria-label="Sobrenome" class="form-control campo-customizado"  onBlur="maiuscula(this);removeAspa(this)" placeholder="Nome"  autocomplete="off" size="30" maxlength="20" value="<%=nomePax%>" required> 
                        </div>
                        <div class="col-md-3 mb-3">
                            <input name="sobrePax<%=p%>"  type="text" id="sobrePax<%=p%>" aria-label="Nome" class="form-control campo-customizado"   onBlur="maiuscula(this);removeAspa(this);" placeholder="Sobrenome"  autocomplete="off"size="30" maxlength="50 " required onChange=document.form1.sobrePax<%=p%>.style.background="#EFF4FA" value="<%=sobrePax%>" > 
                        </div>
                        <div class="col-md-3 mb-3">
                            <input name="tipoPax<%=p%>" type="hidden" value="CPF" id="tipoPax<%=p%>" size="20" maxlength="20"> 
                            <input name="docPax<%=p%>" type="text" value="<%=docPax%>" id="docPax<%=p%>" class="form-control campo-customizado cpf <%if request.cookies("FCNET_MIC")("login") <> "belv" or Request.QueryString("planoId") <> 11760 or Request.QueryString("planoId") = 11761 <> Request.QueryString("planoId") <> 11762 then response.write "valida"%>"  id="docPax<%=p%>" autocomplete="off" size="17" maxlength="11" value="<%=documento%>" placeholder="CPF" aria-label="DocPax" required>
                        </div>
                        <div class="col-md-3 mb-3">
                            <select name="sexoPax<%=p%>" class="form-select campo-customizado" id="sexoPax<%=p%>" required onChange="func_fem(<%=p%>)"  >
                                <option disabled selected value=""  >Sexo</option>
                                <option value="M" <%if sexoPax = "M" then response.Write("selected")%>>Masculino</option>
                                <option value="F" <%if sexoPax = "F" then response.Write("selected")%>>Feminino</option>
                            </select>
                        </div>
                        <div class="col-md-1 mb-3" style="display:flex; justify-content:center;align-items:center">
                            Data de Nascimento:
                        </div>
                        <div class="col-md-1 mb-3">
                            <!-- <input type="date" id="datepicker" class="form-control campo-customizado data" placeholder="Data de Nascimento" maxlength="10" wfd-id="id2"> -->
                            
                            <select name="day<%=p%>" class="form-select campo-customizado" id="day<%=p%>" required onChange="calcage<%=p%>(document.form1.month<%=p%>.value, document.form1.day<%=p%>.value,document.form1.year<%=p%>.value,<%=p%>,<%=n_novo%>)">
                                <option disabled selected value="" <%if dia = 0 then response.Write("selected")%>>Dia</option>
                                <%
                                    aux_dia = 1
                                    while aux_dia < 32
                                        valor_dia = trata_filer(aux_dia,2,"D","0")
                                %>
                                <option value="<%=valor_dia%>" <%if dia = aux_dia then response.Write("selected")%>><%=valor_dia%></option>
                                <%
                                        aux_dia = aux_dia +1
                                    wend
                                %>
                            </select>
                        </div>
                        <div class="col-sm-2 mb-3">
                            <select name="month<%=p%>" class="form-select campo-customizado" id="month<%=p%>" required onChange="calcage<%=p%>(document.form1.month<%=p%>.value, document.form1.day<%=p%>.value,document.form1.year<%=p%>.value,<%=p%>,<%=n_novo%>)">
                                <option disabled selected value="" <%if mes = 0 then response.Write("selected")%>>Mês</option>
                                <option value="01" <%if mes = 1 then response.Write("selected")%>>Janeiro</option>
                                <option value="02" <%if mes = 2 then response.Write("selected")%>>Fevereiro</option>
                                <option value="03" <%if mes = 3 then response.Write("selected")%>>Março</option>
                                <option value="04" <%if mes = 4 then response.Write("selected")%>>Abril</option>
                                <option value="05" <%if mes = 5 then response.Write("selected")%>>Maio</option>
                                <option value="06" <%if mes = 6 then response.Write("selected")%>>Junho</option>
                                <option value="07" <%if mes = 7 then response.Write("selected")%>>Julho</option>
                                <option value="08" <%if mes = 8 then response.Write("selected")%>>Agosto</option>
                                <option value="09" <%if mes = 9 then response.Write("selected")%>>Setembro</option>
                                <option value="10" <%if mes = 10 then response.Write("selected")%>>Outubro</option>
                                <option value="11" <%if mes = 11 then response.Write("selected")%>>Novembro</option>
                                <option value="12" <%if mes = 12 then response.Write("selected")%>>Dezembro</option>
                            </select>
                        </div>
                        <div class="col-md-2 mb-3">
                            <select name="year<%=p%>" class="form-select campo-customizado" id="year<%=p%>" required onChange="calcage<%=p%>(document.form1.month<%=p%>.value, document.form1.day<%=p%>.value,document.form1.year<%=p%>.value,<%=p%>,<%=n_novo%>);verIDade(<%=p%>);">
                                <% 
                                    montaJava = montaJava & "calcage"&p&"(document.form1.month"&p&".value, document.form1.day"&p&".value,document.form1.year"&p&".value,"&p&","&n_novo&");"
                                    if p > n_novo then
                                        response.Write(montarAnosIdoso(ano))
                                    else
                                        response.Write(montarAnosAgora(ano))
                                    end if
                                %>
                            </select>  
                        </div>
                        <div class="col-md-2 mb-3">
                            <input type="text" class="form-control campo-customizado" id="idadePax<%=p%>" name="idadePax<%=p%>" placeholder="Idade" readonly>
                        </div>
                        <div class="col-md-2 mb-3" hidden disabled>
                            <input id="DDDcelularPax<%=p%>" name="DDDcelularPax<%=p%>" type="tel" class="form-control campo-customizado" value="<%=cel%>" size="30" maxlength="38" placeholder="DDD+Celular" onBlur="maiuscula(this);removeAspa(this)">
                        </div>
                        <div class="col-md-2 mb-3" hidden disabled>
                            <input name="emailPax<%=p%>" type="email" class="form-control campo-customizado" id="emailPax<%=p%>" placeholder="teste@teste.com" autocomplete="off" maxlength="100" onBlur="removeAspa(this)" value="<%=emailPax%>" >
                        </div>
                        
                        <!-- estava aqui o gestante -->
                        <%
                            set upRS =objConn.execute("select upgrade.id FROM planos_upgrade_tipo LEFT JOIN planos upgrade on upgrade.up_tipo_id = planos_upgrade_tipo.id where upgrade.id in (SELECT upGradeId from planos_upgrade where planoId = " & planoId & ")")
                            if not upRS.EOF then 
                        %>
                        <div class="col-md-1 mb-3" hidden>
                            <span>Upgrade Cancel:</span>             
                        </div>
                        <div class="col-md-3 mb-3" hidden>
                            <select class="form-select campo-customizado" id="planoCancelamento<%=p%>" name="planoCancelamento<%=p%>" onChange="verUpgrade(<%=p%>)" >
                                <option disabled value="">Selecione um upgrade</option>
                                <option id="0" value="0">Nenhum adicional</option>
                                <%
                                    set rsCancelamento = objConn.Execute("select *, planos.id as idezao from planos inner join valoresdiarios on valoresdiarios.planoId = planos.id where nome like '%cancelamento%' and planos.id NOT IN (SELECT plano_id FROM goAffinity_planos_nao where parceiro_id = "&request.cookies("wlabel")("revId")&") order by planos.id")

                                    while not rsCancelamento.eof
                                        cancelValue = forMoeda(rsCancelamento("preco"),2)
                                        cancelId = rsCancelamento("idezao")                                
                                    %>
                                    <option id="<%=cancelId%>" value="<%=cancelValue%>"<% if cancelId = upgrade_cancel then response.Write("selected")%> ><%= "US$ " & formatnumber(rsCancelamento("i_seg"),2) & " - US$ " & cancelValue%> </option>
                                <%
                                    rsCancelamento.movenext
                                    wend                           
                                %>
                            </select>
                            <input type="hidden" id="planoCancelamento_id<%=p%>" name="planoCancelamento_id<%=p%>" value="0"/>
                        </div>
                        <%  end if
                            set rsCovid = objConn.Execute("select planos.id as idcovid, * from planos inner join planos_upgrade on upGradeId = planos.id where nome like '%covid%' and nacionalidade = '"& planoRS("nacionalidade") & "' and planoId = '"& planoRS("id") & "' and planos.id NOT IN (SELECT plano_id FROM goAffinity_planos_nao where parceiro_id = "&request.cookies("wlabel")("revId")&") order by nacionalidade, planos.id")
                            if not rsCovid.eof then 
                        %>
                        <div class="col-md-1 mb-3">
                            <span>Upgrade Covid-19:</span> 
                        </div>
                        <div class="col-md-3 mb-3">                   
                            <select class="form-select campo-customizado" id="planoCovid<%=p%>" name="planoCovid<%=p%>" onChange="verUpgrade(<%=p%>)">
                                <option disabled value="">Selecione um upgrade</option>
                                <option id="0" value="0">Nenhum adicional</option>    
                                <%                                                   
                                    while not rsCovid.eof
                                        covid_id = rsCovid("idcovid")                                
                                        Call CalcCovid (covid_id, planoId, destino, 0, dias)                                
                                        covidValue = priceCovid
                                %>
                                    <option id="<%=covid_id%>" value="<%=covidValue%>"<% if covid_id = upgrade_covid then response.Write("selected") end if%> >
                                        <%
                                            if planoRS("nacionalidade") = "i" then
                                                Response.write rsCovid("nome") & " - US$ " & forMoeda(covidValue,2)
                                            else
                                                Response.write rsCovid("nome") & " - R$ " & forMoeda(covidValue,2)
                                            end if
                                        %> 
                                    </option>
                                <%
                                    rsCovid.movenext
                                    wend                                                    
                                %>
                            </select>
                            <input type="hidden" id="planoCovid_id<%=p%>" name="planoCovid_id<%=p%>" value="0" />
                        </div>
                        <%  end if %>    
                        <%if planoRS("aceita_gestante") = "S" then%>
                            <div style="display:flex;flex-direction:row">
                            <div class="input-group2" id="gestacao<%=p%>" style="display:none;margin-right:10px;">
                                <div class="input-group-prepend">
                                    <label class="input-group-text" for="inputGroupSelect01">Gestante</label>
                                </div>
                                <select name="gestantePax<%=p%>" class="custom-select  campo" id="gestantePax<%=p%>" onChange="func_gravidez(<%=p%>)" onClick="return verificaGestante(document.getElementById('idadePax<%=p%>').value,document.getElementById('sexoPax<%=p%>').value);"> 
                                    <option value="Nao" <%if gestantePax = "1" then response.Write("selected")%>>Não</option>
                                    <option value="Sim" <%if gestantePax = "2" then response.Write("selected")%>>Sim</option>
                                </select>
                            </div> 
                            <div class="input-group2" id="dados_gestacao<%=p%>" style="display:none">
                                <div class="input-group-prepend">
                                    <label class="input-group-text" for="inputGroupSelect01">Sem. de Gestação</label>
                                </div>
                                <select name="semanasGes<%=p%>" id="semanasGes<%=p%>" class="custom-select campo"  >
                                    <%= montarGestacao(gest) %>
                                </select>
                            </div>
                        </div>

                        <%end if%>  
                    </div>
                    <!-- preco do plano -->
                    <div class="form-group row justify-content-center" style="display:none;">
                    <div class="input-group col-md-3">
                        <%	if planoRS("nacionalidade") = "i" then 
                                typeHtml = "text"
                                tarifa = "Tarifa USD"
                                moeda = "US$"
                        %>
                        <div class="input-group-prepend">
                            <span class="input-group-text">Cambio - R$</span>
                        </div>
                        <% 
                            else 
                                typeHtml = "hidden"
                                tarifa = "Tarifa BRL"
                                moeda = "RS$"
                            end if 
                        %>
                        <input name="cambio<%=p%>" type="<%= typeHtml %>" class="form-control campo " id="cambio<%=p%>" onBlur="maiuscula(this)" value="<%=formoeda(cambio,2)%>" size="10" maxlength="10" readonly>
                    </div>																																		
                    <div class="input-group col-md-3">
                        <div class="input-group-prepend">
                            <span class="input-group-text"><%=tarifa%></span>
                        </div>
                        <%if p <= n_novo then%>
                        <input name="valor_original" type="hidden" id="valor_original<%=p%>" value="<%=forMoeda(tarifapax_original,2)%>">                 
                        <input name="valorUSD<%=p%>" type="text" class="campo form-control" id="valorUSD<%=p%>" onBlur="maiuscula(this)" value="<%=forMoeda(tarifapax,2)%>" maxlength="10" readonly>
                        <input name="valorUSD_Origin<%=p%>" type="hidden" class="campo form-control" id="valorUSD_Origin<%=p%>" onBlur="maiuscula(this)" value="<%if familiar = 1 and p > 1 then response.write forMoeda(0,2) else response.write forMoeda(tarifapax,2) end if%>" maxlength="10">
                        <input name="valorOriginal<%=p%>" type="hidden" class="campo" id="valorOriginal<%=p%>" value="<%=forMoeda(tarifapax_original,2)%>" size="10" maxlength="10" disabled>                
                    </div>
                    <div class="input-group col-md-3">
                        <%	if planoRS("nacionalidade") = "i" then %>
                            <div class="input-group-prepend">
                                <span class="input-group-text">Tarifa BRL</span>
                            </div>
                        <% end if %>
                        <input name="valorBRL<%=p%>" type="<%=typeHtml%>" class="campo form-control" id="valorBRL<%=p%>" onBlur="maiuscula(this)" value="<%=forMoeda(tarifapaxBR,2)%>"  size="10" maxlength="10" readonly> <span class="afterCupom" style="color:red;line-height: 38px;margin-left: 3px;"></span>               
                        <input name="valorBRL_Origin<%=p%>" type="hidden" class="campo form-control" id="valorBRL_Origin<%=p%>" onBlur="maiuscula(this)" value="<%if familiar = 1 and p > 1 then response.write forMoeda(0,2) else response.write forMoeda(tarifapaxBR,2) end if%>" maxlength="10">
                        <%else%>
                        <input name="valor_original" type="hidden" id="valor_original<%=p%>" value="<%=forMoeda(tarifapax_original_idoso,2)%>">
                        <input name="valorUSD<%=p%>" type="text" class="campo form-control" id="valorUSD<%=p%>" onBlur="maiuscula(this)" value="<%=forMoeda(tarifapax_idoso,2)%>" maxlength="10" readonly>
                        <input name="valorUSD_Origin<%=p%>" type="hidden" class="campo form-control" id="valorUSD_Origin<%=p%>" onBlur="maiuscula(this)" value="<%=forMoeda(tarifapax_idoso,2)%>" maxlength="10">
                        <input name="valorOriginal<%=p%>" type="hidden" class="campo" id="valorOriginal<%=p%>" value="<%=forMoeda(tarifapax_original_idoso,2)%>" size="10" maxlength="10" disabled>                
                    </div>
                    <div class="input-group col-md-3">
                        <input name="valorBRL" type="hidden" id="valorBRL" value="<%=forMoeda(tarifapaxBR_idoso,2)%>">
                        <%	if planoRS("nacionalidade") = "i" then %>
                            <div class="input-group-prepend">
                                <span class="input-group-text">Tarifa BRL</span>
                            </div>
                        <% end if %>
                        <input name="valorBRL<%=p%>" type="<%=typeHtml%>" class="campo form-control" id="valorBRL<%=p%>" onBlur="maiuscula(this)" value="<%=forMoeda(tarifapaxBR_idoso,2)%>"  size="10" maxlength="10" readonly><span class="afterCupom" style="color:red;line-height: 38px;margin-left: 3px;"></span>
                        <input name="valorBRL_Origin<%=p%>" type="hidden" class="campo form-control" id="valorBRL_Origin<%=p%>" onBlur="maiuscula(this)" value="<%=forMoeda(tarifapaxBR_idoso,2)%>" maxlength="10" >
                    </div>
                        <%end if%>    
                    
                </div>
            </div>
        </div>
        
        <!-- TUDO QUE TIVER DE SCRIPT VOU COLOCAR AQUI -->
        <script language="JavaScript">
            function changeNation (nac,pax) {
                        
                $("#tipoPax"+pax).val(nac);
                if(nac == 'PASS'){
                    $("#docPax"+pax).attr("placeholder", "ID/PASS");
                    $("#docPax"+pax).removeClass("valida");
                    $("#planoRecep"+pax).removeAttr("hidden");
                    $("#planoRecep"+pax).removeAttr("disabled"); 
                    //$("#planoCep"+pax).removeAttr("hidden");
                    //$("#planoCep"+pax).removeAttr("disabled");
                    $("#origemPass"+pax).removeAttr("hidden", "hidden");
                    $("#origemPass"+pax).removeAttr("disabled", "disabled"); 
                    $('#DDDcelularPax'+pax).unmask();
                    
                        
                }
                if(nac == 'CPF'){
                    $("#docPax"+pax).attr("placeholder", "CPF");
                    $("#docPax"+pax).addClass("valida");
                    $("#planoRecep"+pax).attr("hidden", "hidden");
                    $("#planoRecep"+pax).attr("disabled", "disabled"); 
                    //$("#planoCep"+pax).attr("hidden", "hidden");
                    //$("#planoCep"+pax).attr("disabled", "disabled"); 
                    $("#origemPass"+pax).attr("hidden", "hidden");
                    $("#origemPass"+pax).attr("disabled", "disabled"); 
                    $('#DDDcelularPax'+pax).mask('(00)000000000');
                }
                if (pax == '1' && nac == 'PASS') {
                    $('#cepPax').off('onblur');
                    $(document).ready(function(){

                        $('#cepPax').unmask('00000-000');
                        $('.telefoneP ').unmask('(00)000000000');
                        $('#FoneN').unmask('(00)000000000');  
                    });
                };
                if (pax == '1' && nac == 'CPF')  {
                    $('#cepPax').on('onblur');
                    $(document).ready(function(){

                        $('#cepPax').mask('00000-000');
                        $('.telefoneP ').mask('(00)000000000');
                        $('#FoneN').mask('(00)000000000');  
                    });
                };
                $("#origemPass"+pax).val('');
                $("#docPax"+pax).val('');
                $("#docPax"+pax).focus();

            }
            //nao fazer validacao de cpf em planos receptivos
            
            $("#docPax<%=p%>.valida").change(function(e) { 
                check_cpf_SUSEP( $("#docPax<%=p%>.valida").val(),'docPax<%=p%>.valida','docPax<%=p%>');
            });
            // novo calculo de idade
            function calcage<%=p%>(mm,dd,yy,pax,novo) {
                mm2 = <%=month(date)%>// + 1
                dd2 = <%=day(date)%>
                yy2 = <%=year(date)%>

                if (yy2 < 100) {
                    yy2 = yy2 + 1900
                }
                yourage = yy2 - yy
                if (mm2 == mm) {
                    if (dd2 < dd) {
                        yourage = yourage - 1; 
                    }
                }
                if (mm2 < mm) {
                    yourage = yourage - 1; 
                }
                
                if (yy == "" || mm == "" || dd == ""){
                    document.form1.idadePax<%=p%>.value = 0
                }
                else{
                    document.form1.idadePax<%=p%>.value = yourage
                }      
                if (pax <= novo){
                    if (mm > 0 && dd > 0 && document.form1.idadePax<%=p%>.value == 65){
                        alert('A idade do passageiro <%=p%> não pode ter 65 anos');
                        document.form1.day<%=p%>.value = ""
                        document.form1.month<%=p%>.value = ""
                        document.form1.year<%=p%>.value = ""
                        document.form1.idadePax<%=p%>.value = ""
                        document.form1.year<%=p%>.focus();
                        return false;
                    }
                }
                
                if (pax > novo){
                    if (document.form1.idadePax<%=p%>.value > 85){
                        alert('A idade do passageiro <%=p%> não pode ser maior que 85');
                        document.form1.day<%=p%>.value = ""
                        document.form1.month<%=p%>.value = ""
                        document.form1.year<%=p%>.value = ""
                        document.form1.idadePax<%=p%>.value = ""
                        document.form1.year<%=p%>.focus();
                        return false;
                    }
                    if (document.form1.idadePax<%=p%>.value == 64){
                        alert('A idade do passageiro <%=p%> não pode ser menor que 64');
                        document.form1.day<%=p%>.value = ""
                        document.form1.month<%=p%>.value = ""
                        document.form1.year<%=p%>.value = ""
                        document.form1.idadePax<%=p%>.value = ""
                        document.form1.year<%=p%>.focus();
                        return false;
                    }
                }                                                     
            }
                        
                    
        </script>
        <!-- fimmmmmmm novo campo -->  
        </div>
        <%
            next
        %>                
    </section>

    <!--cupom-->
    <%
        'so mostra cupom para quem tem cadastrado no tavola
        set cpDescRs = objConn.execute("SELECT * FROM cupom_desconto WHERE data_expiracao >= GETDATE() AND cadCliente_id = "& Request.cookies("wlabel")("revId") &" AND ativo = 1")

        if not cpDescRs.eof and blocoId = "" then
    %>
    <section style="text-align:center;">
        <div id="btCupom" class="btn btn-primary" onclick="$('#areaCupom').css('display', 'block');$(this).css('display', 'none')">Possui um cupom de desconto?</div>
        <br>
        <div class="content container sGrid sGrid-pad rounded shadow p-4 mb-4 bg-white" id ="areaCupom"  style="display:none;">
            <div class="formHeader1">
                Cupom <span onclick="$('#areaCupom').css('display', 'none');$('#btCupom').css('display', 'inline-block')" style="float:right;cursor:pointer"><i class="fas fa-times"></i></span>
            </div>
            <br>
            <div id="cupomZone" class="form-group row">
                <div class="input-group" >
                    <br><br><br>
                    <small style="margin:10px auto;">Caso possua um cupom de desconto, insira abaixo: </small>
                    <br><br><br>
                </div>
                <div class="input-group" >
                    <small style="margin:10px auto;color:red;"></small>
                    <br><br><br>
                </div>
                <div class="input-group-prepend col-lg-7" style="margin:0 auto;">
                    <span class="input-group-text">Cupom de desconto</span>
                    <input id="cupom" type="text" maxlength="14"  class="campo form-control" />
                </div>
                <div class="input-group">
                    <div class="btn btn-primary" style="margin:15px auto;" id="cupomBtn" onclick="validaCupom($('#cupom').val())">APLICAR CUPOM</div>
                </div>
            </div>
            <div id="afterDiscount">
            </div>
        </div>
    </section>
    <%
        end if
    %>
            <hr>
        <!-- Início contato de emergência -->

        <div class="accordion" id="accordionExample">
            <div class="accordion-item">
                <h2 class="accordion-header" id="headingOne">
                <button class="accordion-button collapsed" type="button" data-toggle="collapse" data-target="#collapseOne" aria-expanded="false" aria-controls="collapseOne">
                    <i class="fas fa-ambulance"></i> Contato de Emergência
                </button>
                </h2>
                <div id="collapseOne" class="accordion-collapse collapse" aria-labelledby="headingOne" data-parent="#accordionExample">
                <div class="accordion-body">
                    <div class="row">
                        <div class="col-md-4 mb-3"><input name="contatoNome" type="text" class="form-control campo-customizado" id="contatoNome" onBlur="maiuscula(this);removeAspa(this)" size="40" maxlength="100" placeholder="Nome"  autocomplete="off" size="30" maxlength="20" value="<%=contatoNome%>"></div>
                        <div class="col-md-4 mb-3"><input type="text" class="form-control campo-customizado telefone" placeholder="Telefone"></div>
                        <div class="col-md-4 mb-3"><input type="text" class="form-control campo-customizado" placeholder="Endereço"></div>
                    </div>
                </div>
                </div>
            </div>
        </div>

        
        <!-- Fim contato de emergência -->
    <hr>
        <!-- Início dados de pagamento -->
        <h5 class="card-title"><i class="fas fa-dollar-sign"></i> Dados do Comprador</h5><br>
        <div class="row">
            
            <div class="col-md-6 mb-3"><input id="foneN" name="foneN" type="tel" class="telefoneP form-control campo-customizado" value="<%=fone%>" maxlength="15" placeholder="Telefone" autocomplete="off" onBlur="maiuscula(this);removeAspa(this)"></div>
            <div class="col-md-6 mb-3"><input type="email" class="form-control campo-customizado" placeholder="E-mail"></div>
            <div class="col-md-4 mb-3">                    
                <input name="cepPax" type="text" class="form-control campo-customizado" onpaste="return false" id="cepPax" onBlur="pesquisacep(this.value);" value="<%=cep%>" maxlength="10" placeholder="00000-000" required autocomplete="off">
                <a href="http://www.buscacep.correios.com.br/sistemas/buscacep/" target="_blank">Não sabe o CEP? Clique aqui </a>
            </div>
            <div class="col-md-4 mb-3"><input id="enderecoPax" name="enderecoPax" type="text" class="form-control campo-customizado" onBlur="maiuscula(this);removeAspa(this)" value="<%=endereco%>" maxlength="100" placeholder="Rua exemplo" required  autocomplete="off"></div>
            <div class="col-md-2 mb-3"><input name="numeroPax" type="text" class="form-control campo-customizado" id="numeroPax" onBlur="maiuscula(this);removeAspa(this)" value="<%=numero%>" maxlength="30" placeholder="0000" required autocomplete="off" > </div>
            <div class="col-md-2 mb-3"><input name="complementoPax" type="text" class="form-control campo-customizado" id="complementoPax" onBlur="maiuscula(this);removeAspa(this)" value="<%=complemento%>" placeholder="Casa 3/ Bloco 3, Apto 32"  maxlength="30" ></div>
            <div class="col-md-4 mb-3"><input name="bairroPax" type="text" class="form-control campo-customizado" id="bairroPax" onBlur="maiuscula(this);removeAspa(this)" value="<%=bairro%>" required placeholder="..." autocomplete="off" maxlength="30" ></div>
            <div class="col-md-4 mb-3"><input id="cidadePax" name="cidadePax" type="text" class="form-control campo-customizado" onBlur="maiuscula(this);removeAspa(this)" value="<%=cidade%>" required placeholder="..." autocomplete="off"  maxlength="60" ></div>
            <div class="col-md-4 mb-3">
                <select class="form-select campo-customizado" id="ufPax" name="ufPax" required>
                        <option disabled selected value=" " >Selecione:</option>
                        <option value="OU"<%if uf = "OU" then response.write "selected" end if%>>Outro</option>
                        <option value="AC"<%if uf = "AC" then response.write "selected" end if%>>AC</option>
                        <option value="AL"<%if uf = "AL" then response.write "selected" end if%>>AL</option>
                        <option value="AP"<%if uf = "AP" then response.write "selected" end if%>>AP</option>
                        <option value="AM"<%if uf = "AM" then response.write "selected" end if%>>AM</option>
                        <option value="BA"<%if uf = "BA" then response.write "selected" end if%>>BA</option>
                        <option value="CE"<%if uf = "CE" then response.write "selected" end if%>>CE</option>
                        <option value="DF"<%if uf = "DF" then response.write "selected" end if%>>DF</option>
                        <option value="ES"<%if uf = "ES" then response.write "selected" end if%>>ES</option>
                        <option value="GO"<%if uf = "GO" then response.write "selected" end if%>>GO</option>
                        <option value="MA"<%if uf = "MA" then response.write "selected" end if%>>MA</option>
                        <option value="MG"<%if uf = "MG" then response.write "selected" end if%>>MG</option>
                        <option value="MS"<%if uf = "MS" then response.write "selected" end if%>>MS</option>
                        <option value="MT"<%if uf = "MT" then response.write "selected" end if%>>MT</option>
                        <option value="PA"<%if uf = "PA" then response.write "selected" end if%>>PA</option>
                        <option value="PB"<%if uf = "PB" then response.write "selected" end if%>>PB</option>
                        <option value="PE"<%if uf = "PE" then response.write "selected" end if%>>PE</option>
                        <option value="PI"<%if uf = "PI" then response.write "selected" end if%>>PI</option>
                        <option value="PR"<%if uf = "PR" then response.write "selected" end if%>>PR</option>
                        <option value="RN"<%if uf = "RN" then response.write "selected" end if%>>RN</option>
                        <option value="RJ"<%if uf = "RJ" then response.write "selected" end if%>>RJ</option>
                        <option value="RO"<%if uf = "RO" then response.write "selected" end if%>>RO</option>
                        <option value="RR"<%if uf = "RR" then response.write "selected" end if%>>RR</option>
                        <option value="RS"<%if uf = "RS" then response.write "selected" end if%>>RS</option>
                        <option value="SC"<%if uf = "SC" then response.write "selected" end if%>>SC</option>
                        <option value="SE"<%if uf = "SE" then response.write "selected" end if%>>SE</option>
                        <option value="SP"<%if uf = "SP" then response.write "selected" end if%>>SP</option>
                        <option value="TO"<%if uf = "TO" then response.write "selected" end if%>>TO</option>
                    </select>                    
            </div>
        </div>
        <div class="row mt-3">
            <div class="col-md-6 mb-3">
                <div class="card plano shadow">
                    <div class="card-body py-4 px-4">
                        <h5 class="card-title"><i class="fas fa-list"></i> Resumo da Compra</h5><br>
                        <ul>
                            <li>Plano: <%=planoRS("nome")%></li>
                            <li>Origem: Brasil</li>
                            <li>Destino: <%=ViagemRS("nome")%></li>
                            <li>Início da Vigência: <%= diaInicio%> de <%=ConvertMes(mesInicio)%> de <%=anoInicio%></li>
                            <li>Fim da Vigência: <%= diaFim%> de <%=ConvertMes(mesFim)%> de <%=anoFim%></li>
                            <li>Passageiros com até 64 anos: <%=n_novo%></li>
                            <li>Passageiros com mais de 65 anos: <%=n_idoso%></li>
                        </ul>
                    </div>
                </div>
            </div>
            <div class="col-md-6 mb-3">
                <div class="card plano shadow">
                    <div class="card-body py-4 px-4">
                        <h5 class="card-title"><i class="fas fa-shopping-cart"></i> Valor Final</h5><br>
                        <h5 class="preco-plano"><a id="luigi"></a></h5>
                        <span class="preco-por-pessoa"><a id="porPessoa"></a> por pessoa</span><br>
                        <input id="valorParParcelar" name="valorParParcelar"  type="text" class="form-control campo" size="15" maxlength="10" readonly hidden >
                        <b>ou em até <a id="parcelas"></a> sem juros
                        de <a id="valorParcelado"></a></b>
                        <div style="display:none;">
                            <input name="idPlano" type="hidden" id="idPlano" value="<% =planoRS("id") %>">
                            <input name="planoN" type="hidden" id="planoN" value="<% =planoRS("nPlano") %>">
                            <input name="acordo" type="hidden" id="acordo" value="<%=acordo%>">
                            <input name="familiar" type="hidden" id="familiar" value="<%=familiar%>">
                            <input name="planoId" type="hidden" id="planoId" value="<% =planoRS("id") %>">
                            <input name="paxTotal" type="hidden" name="paxTotal" value="<%=nPax%>">
                            <input name="destino" type="hidden" name="destino" value="<%=destino%>">
                            <input name="inicioViagem" type="hidden" id="inicioViagem" value="<%=dataInicio%>">
                            <input name="fimViagem" type="hidden" id="fimViagem" value="<%=dataFim%>">
                            <input name="dias" type="hidden" id="dias" value="<%=DateDiff("d",dataInicio,dataFim)+1%>">
                            <input name="processo" type="hidden" id="processo" value="<%=processo%>">
                            <input name="origemRecep" type="hidden" id="origemRecep" value="<%=origemRecep%>">  
                        </div>
                        <div style="display:none;">
                            <input name="AcaoSubmit" type="hidden" value="0" id="AcaoSubmit">
                            <input name="AcaoSubmitBoleto" type="hidden" value="0" id="AcaoSubmitBoleto">
                            <input name="idVoucherTemp" type="hidden" id="idVoucherTemp" value="<%=idVoucherTemp%>">
                            <input name="carregarReserva" type="hidden" value="<%if request.QueryString("carregarReserva") = 1 then response.Write("1") else response.Write("0")%>" id="carregarReserva">
                        </div>
                        <div class="text-end mb-4 mt-4"><input name="emitir"  type="submit" class="cta" id="submit2" onClick="document.getElementById('AcaoSubmit').value = '0';document.getElementById('AcaoSubmitBoleto').value = '1'; return confere_emissao();" value="EFETUAR PAGAMENTO"></div>
                    </div>

                </div>
            </div>
        </div>
        <!-- Fim dados de pagamento -->
    
    </form>
    </div>
    <footer id="footer">
    <!--#include file ="../Components/Footer.asp"-->
    </footer>

</body>
</html>
<%if(pgtoForma <> "xx" and pgtoForma <> "" and pgtoForma = "CC") or numeroCartao<>"" then%>
<script>cc.style.display = ""; dadosOff.style.display = "";</script>
<%
end if
objConn.close
set objConn = nothing
%>