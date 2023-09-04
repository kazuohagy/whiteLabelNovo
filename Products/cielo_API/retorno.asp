<!--#include file="../../Library/Common/micMainCon.asp" -->
<!--#include file="apiCielo.asp" -->
<%

	processo = request.queryString("pedidoId")
	
	if processo <> "" then
		set tidRS = objConn.execute("SELECT TOP 1 * FROM dados_cielo WHERE pedidoId='"&processo&"' ORDER BY id DESC")
		
		if not tidRS.EOF then
		
		
			retorno_tid = tidRS("tid")
			'valor em centavos
			retorno_valor = tidRS("pedidoValor")
			retorno_moeda = tidRS("pedidoMoeda")
			retorno_data_hora = tidRS("pedidoDataHora")
			retorno_descricao = tidRS("pedidoDescricao")
			retorno_parcelas = tidRS("pagamentoParcelas")
			retorno_status = tidRS("statusCielo")
			retorno_codigo = tidRS("lr")
			
			'Código 	Status 	Meio de pagamento 	Descrição
			'0 	NotFinished 	ALL 	Aguardando atualização de status
			
			'1 	Authorized 		ALL 	Pagamento apto a ser capturado ou definido como pago
			
			'2 	PaymentConfirmed 	ALL 	Pagamento confirmado e finalizado
			
			'3 	Denied 			CC + CD + TF 	Pagamento negado por Autorizador
			
			'10 	Voided 		ALL 	Pagamento cancelado
			
			'11 	Refunded 	CC + CD 	Pagamento cancelado após 23:59 do dia de autorização
			
			'12 	Pending 	ALL 	Aguardando Status de instituição financeira
			
			'13 	Aborted 	ALL 	Pagamento cancelado por falha no processamento ou por ação do AF
			
			'20 	Scheduled 	CC 		Recorrência agendada
			if retorno_status = "1" OR retorno_status = "2" then
				autorizado = "S"
				objConn.Execute("UPDATE emissaoProcesso set pgtoAprovado=2  WHERE id ="&processo)
			else
				autorizado = "N"
				
				Select Case retorno_status
				  Case 0
					status = "N&atilde;o finalizado"
				  Case 2
					status = "N&atilde;o finalizado"
				  Case 3
					status = "Pagamento negado por Autorizador"
				  Case 10
					status = "Pagamento cancelado"
				  Case 11
					status = "Pagamento cancelado ap&oacute;s autoriza&ccedil;&atilde;o"
				  Case 12
					status = "Aguardando Status de institui&ccedil;&atilde;o financeira"
				  Case 13
					status = "Pagamento cancelado"
				  Case 20
					status = "Agendado"
				End Select
				
				objConn.Execute("UPDATE emissaoProcesso set pgtoAprovado=1  WHERE id ="&processo)
			end if
			
			'autorizado
			if autorizado = "S" then
				objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processo&"','Processo aprovado TID: "&tid&" COD: "&autorizacaoArp&" | redirecionado para finalizar')")
				objConn.close
				response.Redirect "../emitir.asp?AcaoSubmit=0&pg=CC&processo="&processo
			'não autorizado
			else
				
				objConn.EXECUTE("INSERT INTO processoHistorico (processoId, obs) VALUES ('"&processo&"','Processo nao autorizado TID: "&tid&" | redirecionado para finalizar')")
				%>
				<html>
					<head>
						<link rel="stylesheet" href="../../CSS/bootstrap/css/bootstrap.min.css">
	
						<style>
							label {
								font-weight: bolder;
							}
							.card{
								margin: 25px;
							}
							#footer {
								position: absolute;
								bottom: 0;
								left: 0;
								width:100%;
							}
						</style>
						<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
						<script src="../JavaScript/jquery-3.5.1.min.js"></script>
						<script src="../CSS/bootstrap/js/bootstrap.min.js"></script>
						<!--#include file="../../Components/HTML_Head3.asp" -->
					</head>
					<body>
						<header>
							<!--#include file ="../../Components/Header.asp"-->
   	 					</header>
						<div class="container">
							<div class="card text-center">
								<h5 class="card-header">
									Transa&ccedil;&atilde;o n&atilde;o autorizada.
								</h5>
								<div class="card-body">
									<div class="card bg-warning">Entre em contato com o emissor do cart&atilde;o para maiores informa&ccedil;&otilde;es.</div><br>
									
									<div class="card text-white bg-danger mb-3">
									  <div class="card-header">RETORNO DA OPERA&Ccedil;&Atilde;O</div>
									  <div class="card-body">
									  
										<h5 class="card-title"><%="C&oacute;digo:"&retorno_codigo&" - "&retorno_descricao%></h5>
										<p class="card-text">
											<%="status:"& retorno_status &" - "&status%>
										</p>
									  </div>
									</div>
									<br>
									<a type="button" class="btn btn-primary" href="<%="../formulario-seguro.asp?orderid="&processo%>">Clique aqui para finalizar esta compra com outro cart&atilde;o</a>
								
								</div>
							</div>
						</div>
						<footer id="footer">
						<!--#include file ="../../Components/Footer.asp"-->
						</footer>
					</body>
				</html>
				<%
				
			end if
	
		else
			response.write("Não há transação para este processo: "& processo)
			response.end()
		end if
	else
		response.write("Não há retorno para o processo: "& processo)
		response.end()
	end if 
%>