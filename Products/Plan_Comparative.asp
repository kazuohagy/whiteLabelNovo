<!--#include file="../Library/Common/micMainCon.asp" -->
<!--#include file="../Library/Common/funcoes.asp" -->
<!--#include file="../Library/Products/montacobertura.asp" -->

<%
    modal = true
    address = "../../Products/Plan_Comparative.asp?"

    set chaveRS = objconn.execute("SELECT NEWID()")
    chave = chaveRS(0)

    cotacao_id = request.QueryString("cotacao_id")
		
	set dados_cotacaoRS = objConn.execute("select cotacao_reg.*, cotacao_reg_pax.* from cotacao_reg inner join cotacao_reg_pax ON cotacao_reg_pax.cotacao_id = cotacao_reg.id INNER JOIN planos ON planos.id = cotacao_reg_pax.plano_id where cotacao_reg.id = "& cotacao_id &" ORDER BY ordemExibicao")    
    

    n_novo = dados_cotacaoRS("nPax_Novo")
    n_idoso = dados_cotacaoRS("nPax_idoso")
    total_pax = dados_cotacaoRS("nPax_total")
    moeda = dados_cotacaoRS("moeda")
    categoria = dados_cotacaoRS("categoria_id")
    upCancel_id = dados_cotacaoRS("upgradeCancel")
    upCovid_id = dados_cotacaoRS("upgradeCovid")
    dim nomeUpCovid
    if upCovid_id <> 0  and dados_cotacaoRS("tarifa_upgradeCovid")  <> 0 then
        set dados_covid = objConn.execute("select * from planos where id ='"&upCovid_id&"'")
        nomeUpCovid = " + "&dados_covid("nome")
    end if
  
    
    function acha_valor_cobertura(plano_id,cob_id)
	set temRS = objConn.execute("SELECT simbolo, valor from coberturasPlanos where planoId in (SELECT top 1 id from planos where nplano in (SELECT nplano from planos where id = '"&plano_id&"') and ageid = 0 order by id) AND coberturaId = '"&cob_id&"' and versao_id = 2")
	    acha_valor_cobertura = temRS(0) & temRS(1)
	    set temRS = nothing	
    end function

    vet_planosTar = ""
    SQL = "SELECT count(id) FROM planos WHERE id in (SELECT plano_id from cotacao_reg_pax where cotacao_id = '"&cotacao_id&"')"
	set cotacaoRS = objConn.execute(SQL)
	n_colunas = cotacaoRS(0)
    set cotacaoRS = nothing
    
    function formataIdioma(txt,idiom)
        select case idiom
            case 2:
            formataIdioma = replace(txt,"SIM","YES")
            formataIdioma = replace(formataIdioma,"NAO","NO")
            formataIdioma = replace(formataIdioma,"NÃO","NO")
            formataIdioma = replace(formataIdioma,"ATE","UNTIL")
            formataIdioma = replace(formataIdioma,"ATÉ","UNTIL")
        case else: formataIdioma=txt
        end select
    end function
    'Funcao para converter o valor da parcela e pegar o numero de parcelas
    
    Function convertParcela(valor)
        parcela = "10x"
        'If valor < 60 then
        '    parcela = "1x"
        'ElseIf valor >= 60 and valor < 90 Then
        '    parcela = "2x"
        'ElseIf valor >= 90 and valor < 120 Then
        '    parcela = "3x"
        'ElseIf valor >= 120 and valor < 150 Then
        '    parcela = "4x"
        'ElseIf valor >= 150 and valor < 1800 Then
        '    parcela = "5x"
        'Else
        '    parcela = "10x"
        'End If
        convertParcela = parcela
    End Function

    Function convertParcelaTotal(valor)
        totalParcelado = valor / 10
        'If valor < 60 then
        '    totalParcelado = valor
        'ElseIf valor >= 60 and valor < 90 Then
        '    totalParcelado = valor / 2
        'ElseIf valor >= 90 and valor < 120 Then
        '    totalParcelado = valor / 3
        'ElseIf valor >= 120 and valor < 150 Then
        '    totalParcelado = valor / 4
        'ElseIf valor >= 150 and valor < 1800 Then
        '    totalParcelado = valor / 5
        'Else
        '    totalParcelado = valor / 10
        'End If
        convertParcelaTotal = totalParcelado 
    End Function 


    %>


<!DOCTYPE html>
<html lang="pt-br">
<head>
    
    <!--#include file="../Components/HTML_Head2.asp" -->
    <style>    
        table {
            width: 100%;
            border-collapse: collapse;
        }

        tbody:nth-child(odd) {
            background: #CCC;
        }

        tbody:hover td[rowspan],
        .hoverzin:hover td {
            background: #ffcccc;
        }

        .table thead th, .newHeader{
            background-color: #292d3c !important ;  
        }

        .card-deck .cardiculo .card{
            margin-right: 5px;
            margin-top: 5px;
        }   

        .card-headers {
            background-color: #292d3c;
            flex: 1 0 100px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .modal-body {

            overflow: auto;
        }
        .precoTo{

            font-size:30px;
            font-weight:400
        }

        .btn-primary{
            color:white;
            background-color:#e94a26;
            border-color:#e94a26;
        }

        .btn-primary.focus,.btn-primary:focus{
            color:white;
            background-color:#c83514;
            border-color:#82230d
        }

        .btn-primary:hover{
            color:white;
            background-color:#c83514;
            border-color:#be3313
        }
        
        .btn-primary.active,.btn-primary:active,.open>.btn-primary.dropdown-toggle{
            color:white;
            background-color:#c83514;
            border-color:#be3313
        }
        
        .btn-primary.active.focus,.btn-primary.active:focus,.btn-primary.active:hover,.btn-primary:active.focus,.btn-primary:active:focus,.btn-primary:active:hover,.open>.btn-primary.dropdown-toggle.focus,.open>.btn-primary.dropdown-toggle:focus,.open>.btn-primary.dropdown-toggle:hover{
            color:white;
            background-color:#a72d11;
            border-color:#82230d
        }

        .btn-primary.active,.btn-primary:active,.open>.btn-primary.dropdown-toggle{
            background-image:none
        }

        .modalHeader{
            padding:9px 15px;
            border-bottom:1px solid #eee;
            background-color: #f78528;
            -webkit-border-top-left-radius: 5px;
            -webkit-border-top-right-radius: 5px;
            -moz-border-radius-topleft: 5px;
            -moz-border-radius-topright: 5px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
        }

        .tev {
            font-family: 'geoMedium';
            border: 2px solid #f78528;
            background-color: white;
            font-size: 16px;
            cursor: pointer;
            border-color: #f78528;
            color: #f78528 !important;
        }

        .tev:hover {
            font-family: 'geoMedium';
            background-color: #f78528;
            color: white !important;
        }

        .icon-span {
            color: #f78528;
        }

        @media screen and (min-width: 800px) {
            .closing {
            padding: 0rem 0rem 1rem !important;
            float: right !important;
            margin-right: -30px !important;
            margin-top: -30px !important;
            background-color: white !important;
            border-radius: 15px !important;
            width: 30px !important;
            height: 30px !important;
            opacity: 1 !important;
            color: black !important;
            line-height: 20px !important;
        }
        }

        @media screen and (max-width: 500px) {
    
            .modal { 
                position: fixed; 
                top: 3%; 
                right: 3%; 
                left: 3%; 
                width: auto; 
                margin: 0; 
            }
            .modal-body { 
                height: 60%; 
            }
            .xizinho{
                font-size: 30px !important;
                color: black !important;
            }
        } 
                
        #valor-original {
            text-decoration: line-through;
            color: #999;
            font-size: 18px;
        }
        /* Estilos do botão */
        .botao-sem-decoracao {
        display: inline-block;
        padding: 0;
        background-color: transparent;
        color: #000;
        text-decoration: none;
        border: none;
        font-size: 16px;
        cursor: pointer;
        }









/* Estilos para o botão de abrir/cerrar */
#headingOne button {
  font-size: 16px;
  font-weight: bold;
  color: #007bff;
}

/* Estilos para o conteúdo do collapsible */
.collapse.show .card-body {
  border: 1px solid #e9ecef;
  padding: 10px;
  border-radius: 5px;
}

/* Estilos para a lista de itens */
.list-group-item {
  border: none;
  padding: 10px;
  margin-bottom: 5px;
}

/* Estilos para os valores riscados */
.text-right del {
  color: red;
  text-decoration: line-through;
}

/* Estilos para os valores atualizados */
.text-right ins {
  color: green;
}

/* Estilos para os valores regulares */
.text-right span {
  color: black;
}

/* Estilos para os cabeçalhos (CÂMBIO DA COTAÇÃO, TARIFA POR PASSAGEIRO, COVID) */
.col.text-center h6 {
  font-size: 16px;
  margin-bottom: 10px;
}

/* Estilos para as divisões entre seções */
.col.text-center hr {
  margin-top: 15px;
  margin-bottom: 15px;
  border: none;
  border-top: 1px solid #ccc;
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
        <section>
            <!--#include file ="../Components/Price_Component.asp"-->
        </section>
        <!-- Progresso da Cotação -->
        <div class="text-center mt-5">
            <h2 id="compra">Passo 1 - Seleção do Plano</h2>
        </div>
        <ul class="progressbar d-flex justify-content-center mb-5">
            <li class="active">Seleção do Plano</li>
            <li>Informar Dados</li>
            <li>Realizar Pagamento</li>
        </ul>
        <!-- fim progresso -->
        <section id="comparativos-section1">
            <div class="container content" id="assiten_collapso">           
                <div class="card-deck">
                <%      
                    dim upCovid_id                
                    planoaux = 0                                          

                    WHILE NOT dados_cotacaoRS.EOF
                        plano_nome = dados_cotacaoRS("plano_nome")
                        plan_id = dados_cotacaoRS("plano_id")
                        upCancel_id = dados_cotacaoRS("upgradeCancel")
                        upCovid_id = dados_cotacaoRS("upgradeCovid")
                        tarifa_upCancel = dados_cotacaoRS("tarifa_upgradeCancel")   
                        tarifa_upCovid = dados_cotacaoRS("tarifa_upgradeCovid")
                        tarifa_upCovid_original = dados_cotacaoRS("tarifa_upgradeCovidOriginal")
                        tarifa_upCancelBR = dados_cotacaoRS("tarifa_upgradeCancel") * cambio
                        tarifa_upCovidBR = dados_cotacaoRS("tarifa_upgradeCovid") * cambio    
                        
                        total = 0
                        totalBR = 0
                        total_original = 0
                        vet_planosTar = vet_planosTar & dados_cotacaoRS("plano_id") & ","	
                                            
                        if n_novo <> 0 then
                            tarifapax = dados_cotacaoRS("tarifa_USD")
                            tarifapaxBR = dados_cotacaoRS("tarifa_BRL")
                            tarifapax_original = dados_cotacaoRS("tarifa_original")

                            For i=0 To n_novo-1
                                tarifapax_temp = dados_cotacaoRS("tarifa_USD")
                                tarifapaxBR_temp = dados_cotacaoRS("tarifa_BRL")
                                                            
                                total = total + tarifapax_temp + tarifa_upCancel + tarifa_upCovid                            
                                totalBR = totalBR + tarifapaxBR_temp + tarifa_upCancelBR + tarifa_upCovidBR                            
                                total_original = total_original + tarifapax_original + tarifa_upCancel + tarifa_upCovid_original

                                dados_cotacaoRS.MOVENEXT                                                                    
                            Next                                     
                        end if

                        if n_idoso <> 0 then
                            tarifapax_idoso = dados_cotacaoRS("tarifa_USD")
                            tarifapaxBR_idoso = dados_cotacaoRS("tarifa_BRL")
                            tarifapax_original_idoso = dados_cotacaoRS("tarifa_original")

                            For i=0 To n_idoso-1
                                tarifapax_temp = dados_cotacaoRS("tarifa_USD")
                                tarifapaxBR_temp = dados_cotacaoRS("tarifa_BRL")

                                total = total + tarifapax_temp + tarifa_upCancel + tarifa_upCovid                            
                                totalBR = totalBR + tarifapaxBR_temp + tarifa_upCancelBR + tarifa_upCovidBR
                                total_original = total_original + tarifapax_original_idoso + tarifa_upCancel + tarifa_upCovid_original

                                dados_cotacaoRS.MOVENEXT
                            Next
                        end if                          
                %>
                    <!-- Inicio dos cards de planos -->
                    <div class="card plano mb-3 mt-5">
                        <div class="card-body px-5 py-5">
                            <div class="row">
                                <div class="col-md-4">
                                    <h5 class="card-title nome-plano"><%=plano_nome & nomeUpCovid%></h5>
                                    Sem franquia ou carência<br>
                                    Plano cobre lazer, estudo ou negócios<br>
                                    Opção de atendimento via WhatsApp<br><br>
                                    <button type="button" class="botao-sem-decoracao" data-toggle="modal" data-target="#modal<%=planoaux%>">Detalhes do plano</button>
                                    <%if upCovid_id <> 0 then 
                                        if tarifa_upCovid  <> 0 then%>
                                        <br>
                                        <button type="button" class="botao-sem-decoracao" data-toggle="modal" data-target="#coronaTable">Visualizar Cobertura Covid</button>
                                    <%  end if
                                    end if %> 
                                </div>
                                <div class="col-md-8">
                                    <div class="row">
                                        <%if moeda = "R$" then %>  
                                        <div class="col-md-5">
                                            <h5 class="preco-plano">R$ <%=formatnumber(total,2)%></h5>
                                            <p><strong id="valor-original">R$ <%=formatnumber(total_original,2)%></strong></p>
                                            <span class="preco-por-pessoa">R$ <%=formatnumber(total/total_pax,2)%> por pessoa</span><br>
                                            <b>ou em até <%convertParcela(formatnumber((total),2))%>  sem juros
                                                de R$ <%=formatnumber((convertParcelaTotal(total)),2)%></b>
                                        </div>
                                        <%else%>
                                        <div class="col-md-5">
                                            <h5 class="preco-plano">R$ <%=formatnumber(totalBR,2)%></h5>
                                            <p><strong id="valor-original">R$ <%=formatnumber(total_original * cambio,2)%></strong></p>
                                            <span class="preco-por-pessoa">R$ <%=formatnumber(totalBR/total_pax,2)%> por pessoa</span><br>
                                            <b>ou em até <%= convertParcela(formatnumber((totalBR),2))%> sem juros
                                                de R$ <%=formatnumber(convertParcelaTotal(totalBR),2)%></b>
                                        </div>
                                        <%end if%>
                                        <div class="col-md-2">
                                            <img src="img/selo-plano.png" width="70px">
                                        </div>
                                        <!-- <button type="button" class="btn btn-primary" data-toggle="modal" data-target="#modal">
                                            <i class="fas fa-chevron-right"></i>
                                        </button> -->
                                        <div class="col-md-5 text-end">
                                            <a href="Product_Form.asp?cotacao_id=<%=cotacao_id%>&planoId=<%=plan_id%>&categoria=<%=categoria%>&chave=<%=chave%>" class="cta">COMPRAR SEGURO</a>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <!-- COMECOUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUU -->
                            <div class="">
                            <div class="text-center" id="headingOne">
                                <h5 class="mb-0 "></h5>
                                    <button name="tarifao"class="btn btn-link " data-toggle="collapse" data-target="#alvo-<%=planoaux%>" aria-expanded="true" aria-controls="alvo-<%=planoaux%>">
                                        Detalhes - Tarifa
                                    </button>
                                </h5>
                            </div>
                            <div id="alvo-<%=planoaux%>" class="collapse indent" data-parent="#assiten_collapso">
                                <div class="card-body ">
                                    
                                    <div class="row">
                                        <div class="col text-center">
                                            CÂMBIO DA COTAÇÃO
                                            <div class="text-right">
                                                <%="R$ " & cambio%>
                                            </div>
                                        </div>
                                        <% if moeda = "R$" then %>
                                        <div class="col text-center">
                                           TARIFA POR PASSAGEIRO
                                            <%
                                            numeracao = 0
                                            For i=0 To n_novo-1
                                            numeracao = numeracao+1
                                            %>  
                                            <div>
                                                <li class="list-group-item">
                                                    Passageiro <%=numeracao%>
                                                    <div style="color: red; list-style:none; text-decoration:line-through;" class="text-right">
                                                        <%
                                                            response.write "R$ " & formatnumber(tarifapax_original ,2)
                                                        %>
                                                    </div>
                                                    <div class="text-right">
                                                        <%
                                                            response.write "R$ " & formatnumber(tarifapax ,2)
                                                        %>
                                                    </div>
                                                </li>
                                            </div>
                                             <% Next 
                                             For i=0 To n_idoso-1
                                            numeracao = numeracao+1                                      
                                            %>
                                                <li class="list-group-item">
                                                Passageiro <%=numeracao%>
                                                    <div style="color: red; list-style:none; text-decoration:line-through;" class="text-right">
                                                <%
                                                    response.write "R$ " & formatnumber(tarifapax_original_idoso ,2)
                                                %>
                                                    </div>
                                                    <div class="text-right">
                                                <%
                                                    response.write "R$ " & formatnumber(tarifapax_idoso,2)
                                                %>
                                                    </div>
                                                </li>
                                            <%Next%>
                                        </div>
                                        <% else %>
                                        <div class="col text-center">
                                           TARIFA POR PASSAGEIRO
                                            <%
                                            numeracao = 0
                                            For i=0 To n_novo-1
                                            numeracao = numeracao+1
                                            %>  
                                            <div>
                                                <li class="list-group-item">
                                                    Passageiro <%=numeracao%>
                                                    <div style="color: red; list-style:none; text-decoration:line-through;" class="text-right">
                                                        <%
                                                            response.write "USD " & formatnumber(tarifapax_original) & " - R$ " & formatnumber(tarifapax_original * cambio,2)
                                                        %>
                                                    </div>
                                                    <div class="text-right">
                                                        <%
                                                            response.write "USD " & formatnumber(tarifapax) & " - R$ " & formatnumber(tarifapax * cambio,2)
                                                        %>
                                                    </div>
                                                </li>
                                            </div>
                                             <% Next 
                                             For i=0 To n_idoso-1
                                            numeracao = numeracao+1                                      
                                            %>
                                                <li class="list-group-item">
                                                Passageiro <%=numeracao%>
                                                    <div style="color: red; list-style:none; text-decoration:line-through;" class="text-right">
                                                <%
                                                    response.write "USD " & formatnumber(tarifapax_original_idoso) & " - R$ " & formatnumber(tarifapax_original_idoso * cambio,2)
                                                %>
                                                    </div>
                                                    <div class="text-right">
                                                <%
                                                    response.write "USD " & formatnumber(tarifapax_idoso) & " - R$ " & formatnumber(tarifapax_idoso * cambio,2)
                                                %>
                                                    </div>
                                                </li>
                                            <%Next%>
                                        </div>  
                                        <%end if%>
                                        <div class="col text-center">
                                            COVID
                                            <%
                                            numeracao = 0
                                            For i=0 To n_novo-1
                                            numeracao = numeracao+1
                                            %>  
                                            <div>
                                                <%
                                                if upCovid_id <> 0 then 
                                                    if tarifa_upCovid  <> 0 then
                                                %>
                                                    <li class="list-group-item">
                                                        Upgrade de Covid-19
                                                        <div style="color: red; list-style:none; text-decoration:line-through;" class="text-right">
                                                            <%="US$ " & formatnumber(tarifa_upCovid_original) & " - R$ " & formatnumber(tarifa_upCovid_original * cambio,2) %>
                                                        </div>
                                                        <div class="text-right">
                                                            <%= "US$" & formatnumber(tarifa_upCovid,2) & " - R$ " & formatnumber(tarifa_upCovid * cambio,2)%>
                                                        </div>
                                                    </li>
                                                <%else%>
                                                    <li class="list-group-item">
                                                        Upgrade de Covid-19
                                                        <div  style="color: red;" class="text-right">
                                                            Upgrade não disponível!
                                                        </div>
                                                    </li>
                                                <%  
                                                    end if
                                                end if 
                                                %>
                                                <!-- aqui copia caso der ruim e coloca em baixo -->
                                            </div>
                                            <% Next 
                                            For i=0 To n_idoso-1
                                            numeracao = numeracao+1%>
                                                <%
                                                if upCovid_id <> 0 then 
                                                    if tarifa_upCovid  <> 0 then
                                                %>
                                                    <li class="list-group-item">
                                                        Upgrade de Covid-19
                                                        <div style="color: red; list-style:none; text-decoration:line-through;" class="text-right">
                                                            <%="US$ " & formatnumber(tarifa_upCovid_original) & " - R$ " & formatnumber(tarifa_upCovid_original * cambio,2) %>
                                                        </div>
                                                        <div class="text-right">
                                                            <%= "US$" & formatnumber(tarifa_upCovid,2) & " - R$ " & formatnumber(tarifa_upCovid * cambio,2)%>
                                                        </div>
                                                    </li>
                                                <%else%>
                                                    <li class="list-group-item">
                                                        Upgrade de Covid-19
                                                        <div  style="color: red;" class="text-right">
                                                            Upgrade não disponível!
                                                        </div>
                                                    </li>
                                                <%  
                                                    end if
                                                end if 
                                                %>
                                            <% Next%>
                                        </div>
                                       
                                    </div>
                                    
                                </div>
                            </div>
                        </div>
                            <!-- TERMINOUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUUU -->
                        </div>
                    </div>
                    <!-- Inicio das modais dos cards -->
                <div class="modal fade" id="modal<%=planoaux%>" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                    <div class="modal-dialog modal-dialog-centered modal-xl">
                        <div class="modal-content">
                            <div class="modal-body">
                                <div class="text-end">
                                    <button type="button" class="btn-close" data-dismiss="modal" aria-label="Close"></button>
                                </div><br>
                                <div class="card plano mb-3">
                                    <div class="card-body px-5 py-5">
                                        <div class="row">
                                            <div class="col-md-4">
                                                <h5 class="card-title nome-plano"><%=plano_nome & nomeUpCovid%></h5>
                                                Sem franquia ou carência<br>
                                                Plano cobre lazer, estudo ou negócios<br>
                                                Opção de atendimento via WhatsApp<br><br>
                                            </div>
                                            <div class="col-md-8">
                                                <div class="row">
                                                <%if moeda = "R$" then %>  
                                                    <div class="col-md-5">
                                                        <h5 class="preco-plano">R$ <%=formatnumber(total,2)%></h5>
                                                        <p><strong id="valor-original">R$ <%=formatnumber(total_original,2)%></strong></p>
                                                        <span class="preco-por-pessoa">R$ <%=formatnumber(total/total_pax,2)%> por pessoa</span><br>
                                                        <b>ou em até <%convertParcela(formatnumber((total),2))%> sem juros
                                                            de R$ <%=formatnumber((convertParcelaTotal(total)),2)%></b>
                                                    </div>
                                                    <%else%>
                                                    <div class="col-md-5">
                                                        <h5 class="preco-plano">R$ <%=formatnumber(totalBR,2)%></h5>
                                                        <p><strong id="valor-original">R$ <%=formatnumber(total_original * cambio,2)%></strong></p>
                                                        <span class="preco-por-pessoa">R$ <%=formatnumber(totalBR/total_pax,2)%> por pessoa</span><br>
                                                        <b>ou em até <%= convertParcela(formatnumber((totalBR),2))%>  sem juros
                                                            de R$ <%=formatnumber(convertParcelaTotal(totalBR),2)%></b>
                                                    </div>
                                                    <%end if%>
                                                    <div class="col-md-2">
                                                        <img src="img/selo-plano.png" width="70px">
                                                    </div>
                                                    <div class="col-md-5 text-end">
                                                        <a href="Product_Form.asp?cotacao_id=<%=cotacao_id%>&planoId=<%=plan_id%>&categoria=<%=categoria%>&chave=<%=chave%>" class="cta">COMPRAR SEGURO</a>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div><br>
                                <h2>Coberturas do Plano</h2><br>
                                <table class="table table-striped">
                                <%

                                SQL = "" & _
                                    " SELECT distinct planos.id, coberturas.ordem, coberturas.id, planos.id, planos.nome, planos.limiteadicional, planos.diaadicional, planos.nacionalidade, planos.versaoTarifa, planos.nPlano," & _							
                                    " A.dias, CASE WHEN nacionalidade = 'i' THEN 'US$' ELSE 'R$' END AS MOEDA, B.Dias AS DiasAnt, B.Preco AS PrecoAnt, " & _														
                                    " A.preco AS preco, A.preco AS precoUS, " & _							
                                    " coberturas.id as coberturaid, coberturas.descritivo as cobertura, " & _
                                    " coberturasplanos.simbolo, " & _
                                    " COALESCE(coberturasplanos.valor,'-') as valorCob, " & _
                                    " planos.ordemExibicao" & _							
                                    " FROM planos " & _							
                                    " INNER JOIN coberturasplanos ON coberturasplanos.planoId = planos.Id " & _
                                    " INNER JOIN coberturas ON coberturasplanos.coberturaId = coberturas.Id  and coberturas.ordem <> 0" & _							
                                    " INNER JOIN valoresdiarios A ON A.planoId = planos.Id AND A.dias = (SELECT MIN(X.dias) FROM valoresdiarios X WHERE X.planoId = planos.Id AND X.Dias >= "&dias&") " & _
                                    " LEFT JOIN valoresdiarios B ON B.planoId = A.planoId AND B.Dias = ( SELECT MAX(X.dias) FROM valoresdiarios X WHERE X.planoId = A.planoId AND X.Dias < A.Dias ) " & _
                                    " WHERE coberturasplanos.versao_id = 2 and  planos.nome = '"&plano_nome&"' and planos.versaoTarifa = '4' and ageid = 0 and planos.id in (SELECT id from planos where ativosite = 'S' and nplano in (SELECT nplano from planos where id in (SELECT plano_id from cotacao_reg_pax where cotacao_id = '"&cotacao_id&"'))) and A.preco IS NOT NULL  " & _														
                                    " ORDER BY  coberturas.ordem, coberturas.id, planos.ordemExibicao"							

                                set cotacaoRS = objConn.execute(SQL)

                                While not cotacaoRS.EOF										 
                                %>
                                    <tr class="hoverzin">
                                    <td><%=cotacaoRS("cobertura")%></td>
                                    <td><%=cotacaoRS("simbolo")%><%=cotacaoRS("valorCob")%> &nbsp;</td>
                                    </tr>
            
                                <%
                                cotacaoRS.MoveNext
                                wend
                                %>
                                <%
                                    set obsRs = objConn.execute("SELECT * from planos where ativosite = 'S' and ageid = 0 and nome = '"&plano_nome&"'")
                                    WHILE NOT obsRs.EOF
                                %>
                                    <tbody>
                                        <tr class="hoverzin">
                                            <td>        
                                                <%=obsRs("nome")%>
                                            </td>
                                            
                                            <td colspan="<%=colunaAux-1%>" align="center"  bgcolor="#FFFFFF" style="color: #333; font-size:11px; border-bottom:#33512F 1px dotted; 
                                                    border-right: #DED0CF 1px solid;">
                                                <%=obsRs("obs")%> &nbsp;
                                            </td>
                                        </tr> 
                                    </tbody>
                                <%
                                    obsRs.MOVENEXT
                                    WEND
                                %>
                                </table>
                            </div>
                        </div>
                    </div>
                </div> 
                <% 
                planoaux = planoaux + 1
                WEND 
                %>
                </div>                                                   
            </div>
        </section>
<%
    if upCovid_id <> 0 then 
         if tarifa_upCovid  <> 0 then
%>
    <section id="comparativos-section3">
        <div id="coronaTable" class="modal fade"  role="dialog" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered modal-xl">
                <div class="modal-content"  style="padding:0px">
                    <div class="text-left">
                        <button type="button" class="btn btn-default btn-lg" data-dismiss="modal">Fechar</button>
                    </div>
                    <div class="modal-body"  style="padding:0px">
                        <div class="table-responsive ">
                            <table class="table table-hover table-bordered" border="0" id="assistencia">                                
                                <thead id="header" style="padding:0px">	
    
                                    <th bgcolor="#f78528" class=" text-white newHeader" scope="col">RESUMO DOS SERVIÇOS DE ASSISTÊNCIA EMERGENCIAL </th> 
                                   
                                    <th bgcolor="#f78528" class="text-white"scope="col">LIMITES DE COBERTURAS </th> 	
                    
                                </thead>
                                    
                                    <%
                                        idv = upCovid_id
                                        idioma = 1

                                        Dim objConf, cobertura, coberturaRS, nPRS2 , nPRS

                                        Set  nPRS  =  objConn.Execute("select nPlano,publicado,ageId,id  from  planos  where id= "&idv )

                                        if nPRS("publicado")= 1 and nPRS("ageId") = 0 then
                                            Set  nPRS2 = objConn.Execute("select  id, obs, eng_obs, nPlano  from  planos  where id= "&idv )

                                        end if
                                    
                                    
                                        if nPRS2.eof then Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where id=  "&idv )
                                       
                                        Set  coberturaRS  =  objConn.Execute("select  distinct coberturas.id as id, coberturas.descritivo as descritivo, coberturas.eng_descritivo as eng_descritivo, coberturasplanos.simbolo as simbolo, coberturasplanos.valor as valor, coberturas.ordem as ordem, coberturas.tipo as tipo  from  coberturas INNER JOIN coberturasplanos ON coberturas.id=coberturasPlanos.coberturaId WHERE coberturasPlanos.versao_id = '2' and   coberturasPlanos.planoId='"&nPRS2("id")&"' GROUP BY  coberturas.id, coberturas.descritivo, coberturas.eng_descritivo, coberturasplanos.simbolo, coberturasplanos.valor, coberturas.ordem, coberturas.tipo ORDER BY coberturas.tipo, coberturas.ordem")

                                        WHILE NOT coberturaRS.EOF

                                            Set  objConf  =  objConn.Execute("select  *  from  coberturasplanos  where  coberturasPlanos.versao_id = '2' and coberturaId='"&coberturaRS("id")&"' AND planoId  ='"&nPRS2("id")&"'")

                                            cor1 = cor1 MOD 2
		                                    if cor1 = 0 then
		                                        BG1="#FFFFFF"
                                            else
                                                BG1="#f4f7fc"
                                            end if

                                            if idioma = 1 then descritivo=coberturaRS("descritivo") else descritivo=coberturaRS("eng_descritivo")
                                            if descritivo = "" or isnull(descritivo) then descritivo=coberturaRS("descritivo")
                                
                                    %>
                                <tbody>
                                    <tr class="hoverzin">
                                        <td>
                                            <%=ucase(descritivo)%>
                                        </td>
                                        <td  scope="row" align="center"  bgcolor="#FFFFFF" style="color:#333; font-size:11px; border-bottom:#33512F 1px dotted; 
                                                border-right: #DED0CF 1px solid;">
                                            <%=objConf("simbolo")%> <%=formataIdioma(objConf("valor"),idioma)%> &nbsp;
                                        </td>
                                    <tr>                                        
                                    <%                                        
                                            coberturaRS.MOVENEXT
                                        WEND                                        
                                    %>
                                </tbody>                                
                            </table>    
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

<%
        end if
    end if
%>

<%
    if upCancel_id <> 0 then 
        if tarifa_upCancel <> 0 then
%>
    <section id="comparativos-section4">
        <div id="cancelTable" class="modal fade"  role="dialog" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered modal-xl">
                <div class="modal-content"  style="padding:0px">
                    <div class="text-left">
                        <button type="button" class="btn btn-default btn-lg" data-dismiss="modal">Fechar</button>
                    </div>
                    <div class="modal-body"  style="padding:0px">
                        <div class="table-responsive ">
                            <table class="table table-hover table-bordered" border="0" id="assistencia">                                
                                <thead id="header" style="padding:0px">	    
                                    <th bgcolor="#f78528" class="text-white newHeader" scope="col">RESUMO DOS SERVIÇOS DE ASSISTÊNCIA EMERGENCIAL </th>                                    
                                    <th bgcolor="#f78528" class="text-white" scope="col">LIMITES DE COBERTURAS </th> 	                    
                                </thead>
                                    
                                    <%
                                        idv = upCancel_id
                                        idioma = 1

                                        Set  nPRS  =  objConn.Execute("select nPlano,publicado,ageId,id  from  planos  where id= "&idv )

                                        if nPRS("publicado")= 1 and nPRS("ageId") = 0 then
                                            Set  nPRS2 = objConn.Execute("select  id, obs, eng_obs, nPlano  from  planos  where id= "&idv )

                                        end if
                                    
                                    
                                        if nPRS2.eof then Set  nPRS2  =  objConn.Execute("select id, obs,ageId,nPlano,eng_obs from  planos  where id=  "&idv )
                                       
                                        Set  coberturaRS  =  objConn.Execute("select  distinct coberturas.id as id, coberturas.descritivo as descritivo, coberturas.eng_descritivo as eng_descritivo, coberturasplanos.simbolo as simbolo, coberturasplanos.valor as valor, coberturas.ordem as ordem, coberturas.tipo as tipo  from  coberturas INNER JOIN coberturasplanos ON coberturas.id=coberturasPlanos.coberturaId WHERE coberturasPlanos.versao_id = '2' and   coberturasPlanos.planoId='"&nPRS2("id")&"' GROUP BY  coberturas.id, coberturas.descritivo, coberturas.eng_descritivo, coberturasplanos.simbolo, coberturasplanos.valor, coberturas.ordem, coberturas.tipo ORDER BY coberturas.tipo, coberturas.ordem")

                                        WHILE NOT coberturaRS.EOF

                                            Set  objConf  =  objConn.Execute("select  *  from  coberturasplanos  where  coberturasPlanos.versao_id = '2' and coberturaId='"&coberturaRS("id")&"' AND planoId  ='"&nPRS2("id")&"'")

                                            cor1 = cor1 MOD 2
		                                    if cor1 = 0 then
		                                        BG1="#FFFFFF"
                                            else
                                                BG1="#f4f7fc"
                                            end if

                                            if idioma = 1 then descritivo=coberturaRS("descritivo") else descritivo=coberturaRS("eng_descritivo")
                                            if descritivo = "" or isnull(descritivo) then descritivo=coberturaRS("descritivo")
                                
                                    %>
                                <tbody>
                                    <tr class="hoverzin">
                                        <td>
                                            <%=ucase(descritivo)%>
                                        </td>
                                        <td  scope="row" align="center"  bgcolor="#FFFFFF" style="color:#333; font-size:11px; border-bottom:#33512F 1px dotted; 
                                                border-right: #DED0CF 1px solid;">
                                            <%=objConf("simbolo")%> <%=formataIdioma(objConf("valor"),idioma)%> &nbsp;
                                        </td>
                                    <tr>                                        
                                    <%                                        
                                            coberturaRS.MOVENEXT
                                        WEND                                        
                                    %>
                                </tbody>                                
                            </table>    
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

<%      
        end if
    end if
%>    
    <section id="comparativos-section2">
        <div id="cobertura_table" class="modal fade"  role="dialog" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered modal-xl">
                <div class="modal-content"  style="padding:0px">
                    <div class="text-left">
                        <button type="button" class="btn btn-default btn-lg" data-dismiss="modal">Fechar</button>
                    </div>
                    <div class="modal-body"  style="padding:0px">
                        <div class="table-responsive ">
                            <table class="table table-hover table-bordered" border="0" id="assistencia">
                                
                                <thead id="header" style="padding:0px">	
                                    <th class="text-white" class="newHeader" scope="col">SERVIÇOS E COBERTURA</th> 
    
                                    <%
                                        dados_cotacaoRS.MOVEFIRST
                                        planoColuna = ""
                                        ip = -1
                                        planos_cotacao = "0"
                                        While NOT dados_cotacaoRS.EOF
                                        planos_cotacao = planos_cotacao & "," & dados_cotacaoRS("plano_id")	
                                            if planoColuna <> dados_cotacaoRS("plano_id") then
                                                            
                                        
                                    %>	
                                            
                                        <th class="text-white"scope="col"> <%=dados_cotacaoRS("plano_nome")%></th> 	
    
                                    <%
                                        end if
                                        planoColuna = dados_cotacaoRS("plano_id")
                                        dados_cotacaoRS.MOVENEXT
                                        Wend
    
                                        SQL = "" & _
                                            " SELECT distinct planos.id, coberturas.ordem, coberturas.id, planos.id, planos.nome, planos.limiteadicional, planos.diaadicional, planos.nacionalidade, planos.versaoTarifa, planos.nPlano," & _							
                                            " A.dias, CASE WHEN nacionalidade = 'i' THEN 'US$' ELSE 'R$' END AS MOEDA, B.Dias AS DiasAnt, B.Preco AS PrecoAnt, " & _														
                                            " A.preco AS preco, A.preco AS precoUS, " & _							
                                            " coberturas.id as coberturaid, coberturas.descritivo as cobertura, " & _
                                            " coberturasplanos.simbolo, " & _
                                            " COALESCE(coberturasplanos.valor,'-') as valorCob, " & _
                                            " planos.ordemExibicao" & _							
                                            " FROM planos " & _							
                                            " INNER JOIN coberturasplanos ON coberturasplanos.planoId = planos.Id " & _
                                            " INNER JOIN coberturas ON coberturasplanos.coberturaId = coberturas.Id  and coberturas.ordem <> 0" & _							
                                            " INNER JOIN valoresdiarios A ON A.planoId = planos.Id AND A.dias = (SELECT MIN(X.dias) FROM valoresdiarios X WHERE X.planoId = planos.Id AND X.Dias >= "&dias&") " & _
                                            " LEFT JOIN valoresdiarios B ON B.planoId = A.planoId AND B.Dias = ( SELECT MAX(X.dias) FROM valoresdiarios X WHERE X.planoId = A.planoId AND X.Dias < A.Dias ) " & _
                                            " WHERE coberturasplanos.versao_id = 2 and planos.versaoTarifa = '4' and ageid = 0 and planos.id in (SELECT id from planos where ativosite = 'S' and nplano in (SELECT nplano from planos where id in (SELECT plano_id from cotacao_reg_pax where cotacao_id = '"&cotacao_id&"'))) and A.preco IS NOT NULL  " & _														
                                            " ORDER BY  coberturas.ordem, coberturas.id, planos.ordemExibicao"							
                                        set cotacaoRS = objConn.execute(SQL)

                                        iC = 0
                                        coberturaIdCtrl = 0
                                        colunaAux = 0

                                        While not cotacaoRS.EOF										
                                            if CINT(coberturaIdCtrl) <> CINT(cotacaoRS("coberturaId")) then
                                
                                    %>
                                
                                </thead>
                                <tbody>
                                    <tr class="hoverzin">
                                        <%=CINT(cotacaoRS("coberturaId"))%>
                                        <td style="color: #00ff5f;">
                                            <%=cotacaoRS("cobertura")%>
                                        </td>
                                        <td scope="row" align="center"  bgcolor="#FFFFFF" style="color: #ff0000; font-size:11px; border-bottom:#33512F 1px dotted; 
                                                border-right: #DED0CF 1px solid;">
                                            <%=cotacaoRS("simbolo")%><%=cotacaoRS("valorCob")%> &nbsp;
                                        </td>
                                        
                                        <%
                                            iC = 1
                                            else
                                        %>
                                        
                                        <td  scope="row" align="center"  bgcolor="#FFFFFF" style="color:#ff00ff; font-size:11px; border-bottom:#33512F 1px dotted; 
                                                border-right: #DED0CF 1px solid;">
                                            <%=cotacaoRS("simbolo")%><%=cotacaoRS("valorCob")%> &nbsp;
                                        </td>
                                        
                                    <%
                                            iC = iC + 1
                                        end if
                                        
                                        if iC = n_colunas then
                                            response.write "</tr>"
                                            iC = 0
                                        end if
                                        
                                        coberturaIdCtrl = cotacaoRS("coberturaId")
                                        colunaAux = colunaAux + 1
                                        cotacaoRS.MOVENEXT
                                        WEND
                                        
                                    %>
                                </tbody>
                                <%
                                    set obsRs = objConn.execute("SELECT * from planos where ativosite = 'S' and ageid = 0 and nplano in (SELECT nplano from planos where id in ("&planos_cotacao&")) order by nPlano")
                                    WHILE NOT obsRs.EOF
                                %>
                                <tbody>
                                    <tr class="hoverzin">
                                        <td>
                                            <%=obsRs("nome")%>
                                        </td>
                                        
                                        <td colspan="<%=colunaAux-1%>" align="center"  bgcolor="#FFFFFF" style="color: #333; font-size:11px; border-bottom:#33512F 1px dotted; 
                                                border-right: #DED0CF 1px solid;">
                                            <%=obsRs("obs")%> &nbsp;
                                        </td>
                                    </tr> 
                                </tbody>
                                <%
                                    obsRs.MOVENEXT
                                    WEND
                                %>
                                
                            </table>        
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>
    <!-- ACABOUUUUUU AQUI -->
    </div>
 
    <footer id = "footer">
    <!--#include file="../Components/Footer.asp" -->
    </footer>

    <script>
            $(document).ready(function () {
      $('.celular').mask('(00) 0000-00000');
      $('.telefone').mask('(00) 0000-00000');
      $('.data').mask('00/00/0000');
      $('.cep').mask('00.000-000');
      $('.cpf').mask('000.000.000-00');
    });
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
    //scroll suave
    $(document).ready(function() {
        // setTimeout(() => {
        $('html, body').animate({
            scrollTop: $('#compra').offset().top - 80
            
        }, 600);
        // }, 1000);
    });
    </script>
</body>
</html>
