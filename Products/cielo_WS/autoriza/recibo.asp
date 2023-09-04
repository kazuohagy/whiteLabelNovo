<!--
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Kit de Integração Cielo
' Versão: 3.0
' Arquivo: consulta_transacao.asp
' Função: Consulta de uma transação na Cielo Ecommerce
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
-->
<%
' Dados obtidos da loja para a transação

' - dados do processo
identificacao = "2682266"
modulo = "CIELO"
operacao = "Consulta"
ambiente = "TESTE"

' - dados do pedido
tid = ""

' Parâmetros
parametros = parametros & "identificacao=" & identificacao
parametros = parametros & "&modulo=" & modulo
parametros = parametros & "&operacao=" & operacao
parametros = parametros & "&ambiente=" & ambiente

parametros = parametros & "&tid=" & tid

' URL de acesso ao Gateway Locaweb
urlLocaWebCE = "http://comercio.locaweb.com.br/comercio.comp"

' Instancia o objeto HttpRequest. 
Set objSrvHTTP = Server.CreateObject("MSXML2.XMLHTTP.3.0") 

' Informe o método e a URL a ser capturada 
objSrvHTTP.open "POST", urlLocawebCE, false 

' Com o método setRequestHeader informamos o cabeçalho HTTP 
objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 

' O método Send envia a solicitação HTTP e exibe o conteúdo da página 
objSrvHTTP.Send(parametros)

' Verificando se a busca foi bem sucedida 
If objSrvHTTP.statusText = "OK" Then

    xml = objSrvHTTP.responseText

    retorno_codigo_erro = pegaValorNode(xml,"erro//codigo")
    retorno_mensagem_erro = pegaValorNode(xml,"mensagem")

    retorno_tid = pegaValorNode(xml,"transacao//tid")
    retorno_pan = pegaValorNode(xml,"transacao//pan")

    retorno_pedido = pegaValorNode(xml,"transacao//dados-pedido//numero")
    retorno_valor = pegaValorNode(xml,"transacao//dados-pedido//valor")
    retorno_moeda = pegaValorNode(xml,"transacao//dados-pedido//moeda")
    retorno_data_hora = pegaValorNode(xml,"transacao//dados-pedido//data-hora")
    retorno_descricao = pegaValorNode(xml,"transacao//dados-pedido//descricao")
    retorno_idioma = pegaValorNode(xml,"transacao//dados-pedido//idioma")

    retorno_produto = pegaValorNode(xml,"transacao//forma-pagamento//produto")
    retorno_parcelas = pegaValorNode(xml,"transacao//forma-pagamento//parcelas")

    retorno_status = pegaValorNode(xml,"transacao//status")

    retorno_codigo_autenticacao = pegaValorNode(xml,"transacao//autenticacao//codigo")
    retorno_mensagem_autenticacao = pegaValorNode(xml,"transacao//autenticacao//mensagem")
    retorno_data_hora_autenticacao = pegaValorNode(xml,"transacao//autenticacao//data-hora")
    retorno_valor_autenticacao = pegaValorNode(xml,"transacao//autenticacao//valor")
    retorno_eci_autenticacao = pegaValorNode(xml,"transacao//autenticacao//eci")

    retorno_codigo_autorizacao = pegaValorNode(xml,"transacao//autorizacao//codigo")
    retorno_mensagem_autorizacao = pegaValorNode(xml,"transacao//autorizacao//mensagem")
    retorno_data_hora_autorizacao = pegaValorNode(xml,"transacao//autorizacao//data-hora")
    retorno_valor_autorizacao = pegaValorNode(xml,"transacao//autorizacao//valor")
    retorno_lr_autorizacao = pegaValorNode(xml,"transacao//autorizacao//lr")
    retorno_arp_autorizacao = pegaValorNode(xml,"transacao//autorizacao//arp")

    retorno_codigo_cancelamento = pegaValorNode(xml,"transacao//cancelamento//codigo")
    retorno_mensagem_cancelamento = pegaValorNode(xml,"transacao//cancelamento//mensagem")
    retorno_data_hora_cancelamento = pegaValorNode(xml,"transacao//cancelamento//data-hora")
    retorno_valor_cancelamento = pegaValorNode(xml,"transacao//cancelamento//valor")

    retorno_codigo_captura = pegaValorNode(xml,"transacao//captura//codigo")
    retorno_mensagem_captura = pegaValorNode(xml,"transacao//captura//mensagem")
    retorno_data_hora_captura = pegaValorNode(xml,"transacao//captura//data-hora")
    retorno_valor_captura = pegaValorNode(xml,"transacao//captura//valor")    

    retorno_url_autenticacao = pegaValorNode(xml,"transacao//url-autenticacao")

    ' Se não ocorreu erro exibe parâmetros
    If retorno_codigo_erro = "" Then
        Response.write "<b> TRANSAÇÃO </b><br>"
        Response.write "<b>Código de identificação do pedido (TID): </b>" & retorno_tid & "<br>" 
        Response.write "<b>PAN do pedido (pan): </b>" & retorno_pan & "<br>" 
        
        Response.write "<b>Número do pedido (numero): </b>" & retorno_pedido & "<br>"
        Response.write "<b>Valor do pedido (valor): </b>" & retorno_valor & "<br>"
        Response.write "<b>Moeda do pedido (moeda): </b>" & retorno_moeda & "<br>" 
        Response.write "<b>Data e hora do pedido (data-hora): </b>" & retorno_data_hora & "<br>"
        Response.write "<b>Descrição do pedido (descricao): </b>" & retorno_descricao & "<br>"
        Response.write "<b>Idioma do pedido (idioma): </b>" & retorno_idioma & "<br>"

        Response.write "<b>Forma de pagamento (produto): </b>" & retorno_produto & "<br>"
        Response.write "<b>Número de parcelas (parcelas): </b>" & retorno_parcelas & "<br>"

        Response.write "<b>Status do pedido (status): </b>" & retorno_status & "<br>"

        Response.write "<b>URL para autenticação (url-autenticacao): </b>" & retorno_url_autenticacao & "<br><br>"

        Response.write "<b> AUTENTICAÇÃO </b><br>"
        Response.write "<b>Código da autenticação (codigo): </b>" & retorno_codigo_autenticacao & "<br>"
        Response.write "<b>Mensagem da autenticação (mensagem): </b>" & retorno_mensagem_autenticacao & "<br>"
        Response.write "<b>Data e hora da autenticação (data-hora): </b>" & retorno_data_hora_autenticacao & "<br>"
        Response.write "<b>Valor da autenticação (valor): </b>" & retorno_valor_autenticacao & "<br>" 
        Response.write "<b>ECI da autenticação (eci): </b>" & retorno_eci_autenticacao & "<br><br>"

        Response.write "<b> AUTORIZAÇÃO </b><br>"
        Response.write "<b>Código da autorização (codigo): </b>" & retorno_codigo_autorizacao & "<br>"
        Response.write "<b>Mensagem da autorização (mensagem): </b>" & retorno_mensagem_autorizacao & "<br>"
        Response.write "<b>Data e hora da autorização (data-hora): </b>" & retorno_data_hora_autorizacao & "<br>"
        Response.write "<b>Valor da autorização (valor): </b>" & retorno_valor_autorizacao & "<br>" 
        Response.write "<b>LR da autorização (LR): </b>" & retorno_lr_autorizacao & "<br>"
        Response.write "<b>ARP da autorização (ARP): </b>" & retorno_arp_autorizacao & "<br><br>"

        Response.write "<b> CAPTURA </b><br>"
        Response.write "<b>Código do captura (codigo): </b>" & retorno_codigo_captura & "<br>"
        Response.write "<b>Mensagem do captura (mensagem): </b>" & retorno_mensagem_captura & "<br>"
        Response.write "<b>Data e hora do captura (data-hora): </b>" & retorno_data_hora_captura & "<br>"
        Response.write "<b>Valor do captura (valor): </b>" & retorno_valor_captura & "<br><br>"

        Response.write "<b> CANCELAMENTO </b><br>"
        Response.write "<b>Código do cancelamento (codigo): </b>" & retorno_codigo_cancelamento & "<br>"
        Response.write "<b>Mensagem do cancelamento (mensagem): </b>" & retorno_mensagem_cancelamento & "<br>"
        Response.write "<b>Data e hora do cancelamento (data-hora): </b>" & retorno_data_hora_cancelamento & "<br>"
        Response.write "<b>Valor do cancelamento (valor): </b>" & retorno_valor_cancelamento & "<br>"
    Else
        Response.write "<b>Erro: </b>" & retorno_codigo_erro & "<br>" 
        Response.write "<b>Mensagem: </b>" & retorno_mensagem_erro & "<br>" 
    End If		

End If

Set objSrvHTTP = Nothing 

' ################################################################################################
' pegaValorNode
' Retorno o valor específico de um Node de um XML
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
%>