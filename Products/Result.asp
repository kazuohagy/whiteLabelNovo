<!--#include file="../../Library/Common/micMainCon.asp" -->
<%
voucher = request.QueryString("voucher")
processo = request.QueryString("processo")
if processo <> "" then
    set processoRS = objConn.execute("SELECT * FROM voucher WHERE processoid ='"&processo&"'")
    set processoRS2 = objConn.execute("SELECT planos.nome as nome_plano,destino,vd.nome as destino_nome, * FROM voucher INNER JOIN planos ON voucher.planoId = planos.id INNER JOIN viagem_destino as vd ON voucher.destino = vd.id WHERE processoid ='"&processo&"'")
    set processoRS3 = objConn.execute("SELECT * FROM emissaoProcesso where id ='"&processo&"'")

    dim mesInicio, mesFim, diaInicio, diaFim, anoInicio, anoFim
    'Recuperando as datas de inicio da viagem'
    diaInicio = LEFT(processoRS("inicioVigencia"),2)
    mesInicio = MID(processoRS("inicioVigencia"),4,2)
    anoInicio = RIGHT(processoRS("inicioVigencia"),4)

    'Recuperando as datas de fim da viagem'		
    diaFim = LEFT(processoRS("fimVigencia"),2)
    mesFim = MID(processoRS("fimVigencia"),4,2)
    anoFim = RIGHT(processoRS("fimVigencia"),4)

  'Converte mes em numero para nome do mes em ASP CLASSIC
  Function ConvertMes(mes)
      Select Case mes
          Case "01": ConvertMes = "Janeiro"
          Case "02": ConvertMes = "Fevereiro"
          Case "03": ConvertMes = "Mar&ccedil;o"
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
End if
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Compra Finalizada!</title>
    <!--#include file="../Components/HTML_Head2.asp" -->   
    <style>
               body {
            font-family: Arial, sans-serif;
        }
        .table-container {
            max-width: 800px;
            margin: 20px auto;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        .table {
            width: 100%;
            border-collapse: collapse;
            background-color: #343a40;
            border-radius: 5px;
            overflow: hidden;
        }
        .table thead {
            background-color: #212529;
            color: #fff;
        }
        .table th, .table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #454d55;
        }
        .table tr:last-child td {
            border-bottom: none;
        }
        .table tr:hover {
            background-color: #454d55;
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
            <h2>Finalização</h2>
            <p>Sua compra foi finalizada com sucesso. Abaixo as informações sobre sua compra.</p><br>
            <button onClick="abrirLinksEmNovaAba()" class="cta" style="display:none">BAIXAR CERTIFICADO</button>
            
        <table width="100%" class="table" style="border-radius:20px; border: 2px solid black; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
            <thead>
                <tr> 
                    <th><i class="fas fa-user-alt"></i> Nome</th>
                    <th><i class="fas fa-list"></i> Numero do Voucher</th>
                    <th><i class="fas fa-download"></i> Voucher baixar</th>
                </tr>
            </thead>

            <tbody style="padding: 10px;">
            <%
            nPaxNovo = 0
            nPaxIdoso = 0
            While Not processoRS.EOF		
            %>
                <tr class="linhaCotacao" id="linhaCotacao">
                    <td><%=processoRS("nome")%></td>
                    <td><%=processoRS("voucher")%></td>
                    <td><a href="../Library/Certificates/geraCertificado.asp?idioma=1&mobile=S&voucher=<%=processoRS("voucher")%>"  class="cta baixar">BAIXAR CERTIFICADO</a></td>
                </tr>
                <% if processoRS("idade") < 65 then
                    nPaxNovo = nPaxNovo + 1
                else
                    nPaxIdoso = nPaxIdoso + 1
                end if
                %>
            <%
                processoRS.MoveNext
                Wend
            %>
        </table>

        </div><br>
        <div class="row mt-3">
            <div class="col-md-6 mb-3">
                <div class="card plano shadow">
                    <div class="card-body py-4 px-4">
                        <h5 class="card-title"><i class="fas fa-list"></i> Resumo da Compra</h5><br>
                        <ul>
                            <li>Plano: <%=processoRS2("nome_plano")%></li>
                            <li>Origem: Brasil</li>
                            <li>Destino: <%=processoRS2("destino_nome")%></li>
                            <li>Início da Vigência: <%=diaInicio%> de <%=convertMes(mesInicio)%> de <%=anoInicio%></li>
                            <li>Fim da Vigência: <%=diaFim%> de <%=convertMes(mesFim)%> de <%=anoFim%></li>
                            <li>Passageiros com até 64 anos: <%=nPaxNovo%></li>
                            <li>Passageiros com mais de 65 anos: <%=nPaxIdoso%></li>
                        </ul>
                    </div>
                </div>
            </div>
            <div class="col-md-6 mb-3">
                <div class="card plano shadow">
                    <div class="card-body py-4 px-4">
                        <h5 class="card-title"><i class="fas fa-shopping-cart"></i> Valor Final</h5><br>
                        <h5 class="preco-plano">R$ <%=processoRS2("totalBRL")%></h5>
                        <span class="preco-por-pessoa">R$ <%=formatnumber(processoRS2("totalBRL")/processoRS3("nPax"))%> por pessoa</span><br>
                        <b> <%=processoRS3("parcelas")%>x sem juros
                        de R$ <%=formatnumber(processoRS2("totalBRL")/processoRS3("parcelas"))%></b>
                        <div class="text-end mb-4 mt-5">&nbsp;</div>  
                    </div>
                </div>
            </div>
        </div>

    </div>
    <footer id ='footer'>
    <!--#include file ="../Components/Footer.asp"-->
    </footer>
    <script>
        // Função para baixar todas
		function abrirLinksEmNovaAba() {
			var btn = document.querySelectorAll(".baixar");

            btn.forEach(function(botao) {
                var url = botao.getAttribute("href");
                window.open(url, "_blank");
            });
		}
    </script>
</body>
</html>
