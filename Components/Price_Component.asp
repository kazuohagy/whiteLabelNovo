<!--O endereço do form é decidido pela variável address, e se vai ser um modal ele é decidido pela var modal-->
<%  
    Response.ContentType = "text/html"
    Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
    Response.CodePage = 65001
    Response.CharSet = "UTF-8"

	cotacao_id = request.QueryString("cotacao_id")
    plano_id = request.QueryString("planoId")
    categoria_id = request.QueryString("categoria")
    upgrade_covid = 0
    familiar = request.QueryString("familiar")  
    dias = 0
    nPax_novo = ""
    nPax_idoso = ""

    set rsCategoria = objConn.Execute("select * from categoria where ativo = 1 order by nome")

    if plano_id <> "" then
        set rsDestino = objConn.Execute("select viagem_destino.nome as destinoNome, destinoId from viagem_destino inner join viagem_destinoPlano on viagem_destino.id = viagem_destinoPlano.destinoId inner join planos on viagem_destinoPlano.plano = planos.nPlano where ativo = 1 and planos.id = "& plano_id &" order by viagem_destino.ordem")
    else
        set rsDestino = objConn.Execute("select id as destinoId, nome as destinoNome from viagem_destino where ativo = 1 order by ordem")
    end if
	
    set rsCancelamento = objConn.Execute("select * from planos where nome like '%cancelamento%' and up_tipo_id = 3 and publicado = 1 and id NOT IN (SELECT plano_id FROM goAffinity_planos_nao where parceiro_id = "&request.cookies("wlabel")("revId")&") order by id")

    set rsCovid = objConn.Execute("select * from planos where nome like '%covid%' and up_tipo_id = 8 and publicado = 1 and id NOT IN (SELECT plano_id FROM goAffinity_planos_nao where parceiro_id = "&request.cookies("wlabel")("revId")&") order by nacionalidade, id")
	
	set cambioRS = objConn.execute("select top 1 usdMic from cadCambio where data <= GETDATE() order by data desc")
	cambio = cambioRS("usdMic")
	
	set rsPaises = objConn.execute("select * from paisesLista order by id")
	
	if cotacao_id <> "" then
        set dados_cotacaoRS = objConn.execute("select cotacao_reg.*, cotacao_reg_pax.* from cotacao_reg inner join cotacao_reg_pax ON cotacao_reg_pax.cotacao_id = cotacao_reg.id INNER JOIN planos ON planos.id = cotacao_reg_pax.plano_id where cotacao_reg.id = "& cotacao_id &" and reg_data > CAST(GETDATE() AS DATE) ORDER BY ordemExibicao") 

        if dados_cotacaoRS.eof then
            response.write "<script>alert('\Use um Registro do dia de hoje!')</script>"
            response.write "<script>window.history.back()</script>"
            objConn.close
            response.End()
        end if

        categoria_id = dados_cotacaoRS("categoria_id")
        destino_id = dados_cotacaoRS("destino_id")
        nPax = dados_cotacaoRS("nPax_total")
        nPax_novo = dados_cotacaoRS("nPax_Novo")
        nPax_idoso = dados_cotacaoRS("nPax_Idoso")
        dias = dados_cotacaoRS("vigencia")
        data_inicio = cdate(dados_cotacaoRS("data_inicio"))
        data_fim = cdate(dados_cotacaoRS("data_fim"))
        familiar = dados_cotacaoRS("familiar")
        upgrade_cancel = dados_cotacaoRS("upgradeCancel")
        upgrade_covid = dados_cotacaoRS("upgradeCovid")
	end if
%>
<style>
    .calendar-span {
        float: right;
        margin-right: 10px;
        margin-top: -30px;
        position: relative;
        z-index: 2;       
        cursor: pointer; 
    }

    .comparativo-opcoes p {
        font-size: 14px !important;
    }

    #cambio-label {
        font-size: 17px !important;
    }
    .btn-orange {
        color: #ffffff;
        background-color: #ad0000;
    }

    .btn {
        border-radius: 0%;
    }

    .datepicker th {
        /* back do topo*/
        background: #292d3c;
        color: white;
        padding: 10px;
    }

    .datepicker td {
        padding: 10px 15px;
    }

    .form-control, .custom-select {
        cursor: pointer;
    }

    .custom-control label {
        font-size: 0.9rem;
    }      
    
    #comparativo-box #comparativo-title {
        -webkit-transition: all .3s ease; /* Safari and Chrome */
        -moz-transition: all .3s ease; /* Firefox */
        -o-transition: all .3s ease; /* IE 9 */
        -ms-transition: all .3s ease; /* Opera */
        transition: all .3s ease;
        position:relative;
    }

    #comparativo-box:hover #comparativo-title {
        -webkit-backface-visibility: hidden;
        backface-visibility: hidden;
        -webkit-transform:translateZ(0) scale(1.20); /* Safari and Chrome */
        -moz-transform:scale(1.20); /* Firefox */
        -ms-transform:scale(1.20); /* IE 9 */
        -o-transform:translatZ(0) scale(1.20); /* Opera */
        transform:translatZ(0) scale(1.20);        
    }

    @media (max-width: 1200px) {
        #switchrow {
            max-width: 70%;
        }
    }
.col-md-3 {
  position: relative;
}

.input-container {
  position: relative;

}
.recomendado {
    background: #00213d;
    border-radius: 8px;
    display: flex;
    align-items: center;
    text-align: center;
    text-transform: uppercase;
    color: #ffffff;
    font-size: 0.8em;
    padding: 0 15px;
    width: fit-content;
    position: absolute;
    top: -10px;
    left: 50%;
    transform: translateX(-50%);
}
 .options{
  display: flex;
}
 .options{
  display: none;
  position: absolute;
  top: 55px;
  left: 0;
  border: 1px solid #00213D;
  border-radius: 18px;
  background: #fff;
  height: 100px;
  z-index: 100;
}
</style>
<link rel="stylesheet" href="../CSS/bootstrap-datepicker/bootstrap-datepicker.min.css">
<script src="../CSS/bootstrap-datepicker/bootstrap-datepicker.min.js"></script>
<script src="../CSS/bootstrap-datepicker/bootstrap-datepicker.pt-BR.min.js"></script>
<script>
    $(document).ready(function () {
		<%if not rsCancelamento.eof then%>
        $(".upCancel").change(verificaCancel);
		<%end if
		  if not rsCovid.eof then%>
        $(".upCovid").change(verificaCovid);
		<%end if%>
        $(".verificaFamiliar").change(verificaFamiliar);
		$(".verificaEstudante").change(verificaEstudante);
        $(".verificaReceptivo").change(verificaReceptivo);

        $(".calendar").datepicker({
            "language": "pt-BR",
            "startDate": "+0d",
            "endDate": "+5y",
            "maxViewMode": "2",
            "autoclose": true,
            "keyboardNavigation": false
        });

        $("#inicioViagem").datepicker().on("changeDate", function () {
            $("#fimViagem").datepicker("setStartDate", $("#inicioViagem").val());
            if($("#fimViagem").val() != "") {
                let cOne = $("#inicioViagem").datepicker("getDate");
                let cTwo = $("#fimViagem").datepicker("getDate");
                let qntDias = Math.round((cTwo - cOne) / (1000 * 3600 * 24) + 1);
                //escreve a quantidade de dias
                $("#days").html(qntDias + " Dias");
            } else {
                $("#days").html(0 + " Dias");
            }
        });

        $("#fimViagem").datepicker().on("changeDate", function () {
            $("#inicioViagem").datepicker("setEndDate", $("#fimViagem").val());
            let cOne = $("#inicioViagem").datepicker("getDate");
            let cTwo = $("#fimViagem").datepicker("getDate");
            let qntDias = Math.round((cTwo - cOne) / (1000 * 3600 * 24) + 1);
            //escreve a quantidade de dias
            $("#days").html(qntDias + " Dias");
        });;

        function calculateDateDiff(endDate, startDate) {
            if (endDate && startDate) {
                var start = new Date(swapDayMonth(startDate));
                var end = new Date(swapDayMonth(endDate));
                return parseInt((end.getTime() - start.getTime()) / (24 * 3600 * 1000) + 1);
            }
             return 0;

        }

        function swapDayMonth(date) {
            var day = date.substr(0, 2);
            var month = date.substr(3, 2);
            var rest = date.substr(5);

            return month + "/" + day + "/" + rest;
        }

        function verificaFamiliar() {            
            if($("#idadeMaior").val() > 0) {                
                $("#familiar")[0].checked = false
            }            
        }
		
		 
        function verificaEstudante() {         
            if($("#idadeMaior").val() > 0 && ($("#categoria").val == "11" || document.getElementById('categoria').value == "11") ) {                
                document.getElementById("categoria").selectedIndex = "0";
            }            
        }

		function verificaReceptivo() {    
            console.log("ola mario");   
            console.log("AAAAAAA" + $('#categoria').val());
            var cat = $('#categoria').parent();
            var des = $('#destino').parent();  
            if($("#categoria").val == "7" || document.getElementById('categoria').value == "7") {    
                $("#planoReceptivo").removeAttr("hidden");
                $("#planoReceptivo").removeAttr("disabled");      
                $("#planoPaises").removeAttr("hidden");
                $("#planoPaises").removeAttr("disabled");
                cat.removeClass("col-md-6"); 
                des.removeClass("col-md-6");
                cat.addClass("col-md-4"); 
                des.addClass("col-md-4"); 
                $("#planoPaises").attr("required", "required");
            }
            else {
                $("#planoReceptivo").attr("hidden", "hidden");
                $("#planoReceptivo").attr("disabled", "disabled");         
                $("#planoPaises").attr("hidden", "hidden");
                $("#planoPaises").attr("disabled", "disabled");
                $("#planoPaises").removeAttr("required");
                des.removeClass("col-md-4");
                cat.removeClass("col-md-4");
                des.addClass("col-md-6"); 
                cat.addClass("col-md-6"); 
            }        
        }

		<%if not rsCancelamento.eof then%>
        function verificaCancel() {            
            if ($("#upgradeCancelamento")[0].checked == true && $("#familiar")[0].checked == false && $("#destino").val() != 4) {
                $("#planoCancelamento").removeAttr("hidden");
                $("#planoCancelamento").removeAttr("disabled");
                $("#planoCancelamento").attr("required", "required");
            }
            else {
                $("#planoCancelamento").attr("hidden", "hidden");
                $("#planoCancelamento").attr("disabled", "disabled");
                $("#planoCancelamento").removeAttr("required");
                $("#upgradeCancelamento")[0].checked = false
            }
        }
		<%end if%>
        
		<%if not rsCovid.eof then%>
			var selected_covid = <%=upgrade_covid%>;

			var list_covid_I = {   
				<%                                 
					While rsCovid("nacionalidade") = "i"            
						list_covid = list_covid & Chr(34) & rsCovid("nome") & Chr(34) & ":" &  Chr(34) & rsCovid("id") &  Chr(34) & ", "
					rsCovid.movenext
					wend
					
					length = len(list_covid) -2                
					Response.write left (list_covid,length)
				%>
			};

			var list_covid_N = {            
				<%
					list_covid = ""

					While not rsCovid.eof            
						list_covid = list_covid & Chr(34) & rsCovid("nome") & Chr(34) & ":" &  Chr(34) & rsCovid("id") &  Chr(34) & ", "
					rsCovid.movenext
					wend
					rsCovid.movefirst

					length = len(list_covid) -2                
					Response.write left (list_covid,length)
				%>    
			};

			function verificaCovid() {
				var days = calculateDateDiff($("#inicioViagem").val(), "<%=Date%>");
				if ($("#upgradeCovid")[0].checked == true && $("#familiar")[0].checked == false) {
					$("#planoCovid").removeAttr("hidden");
					$("#planoCovid").removeAttr("disabled");
					$("#planoCovid").attr("required", "required");
					if ($("#planoCovid option:selected").val()) {
						selected_covid = $("#planoCovid option:selected").val();
					}                
					$("#planoCovid").empty();                

					if ($("#destino").val() != 4) {                    
						appendCovid(list_covid_I);
					}
					else {
						appendCovid(list_covid_N);
					}
				}
				else {
					$("#planoCovid").attr("hidden", "hidden");
					$("#planoCovid").attr("disabled", "disabled");
					$("#planoCovid").removeAttr("required");
					$("#upgradeCovid")[0].checked = false
				}
			}

			function appendCovid(list_covid) {            
				$('#planoCovid').append("<option disabled selected value=''>Selecione um upgrade</option>");
				$.each( list_covid, function( key, value ) {
					$('#planoCovid').append($('<option>', { value : value }).text(key));            

					if ($("option[value="+value+"]").val() == selected_covid) {
						$("option[value="+value+"]").attr("selected", "selected");
					}                
				});
			}
		<%end if%>

        <%if cotacao_id <> "" then %>
            $("#inicioViagem").datepicker("setEndDate", $("#fimViagem").val());
            $("#fimViagem").datepicker("setStartDate", $("#inicioViagem").val());
			<%if not rsCancelamento.eof then%>
            verificaCancel();
			<%end if
			  if not rsCovid.eof then%>
            verificaCovid();          
			<%end if%>
        <% end if%>
    });
</script>


    <div class="card formulario-cotacao shadow">
        <form id="price_component" class="mx-4 my-2" onSubmit="return verificaCotacao();"action="../Library/Products/setup_init.asp">
     
     


        <!-- Modal -->
        <!--Preencha primeiro a data de embarque-->
        <div class="modal fade" id="modalEmbarque" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="exampleModalLabel">Preencha a data de embarque</h5>
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                
                </div>
            </div>
        </div>
        <!--Preencha primeiro a data de desembarque-->
        <div class="modal fade" id="modalDesembarque" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="exampleModalLabel">Preencha a data de desembarque</h5>
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                
                </div>
            </div>
        </div>
        <!--Preencha primeiro os passageiros-->
        <div class="modal fade" id="modalPassageiroJ" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="exampleModalLabel">Preencha a quantidade de passageiros Jovens</h5>
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                
                </div>
            </div>
        </div> 
        <!--Preencha primeiro os passageiros-->
        <div class="modal fade" id="modalPassageiroI" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="exampleModalLabel">Preencha a quantidade de passageiros Idosos</h5>
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                
                </div>
            </div>
        </div>
        <div class="card-body px-5 py-5">
            <h2 class="titulo col-12 mb-5">
            Faça agora uma cotação do seu <span style="font-weight: 600; font-style: italic;">seguro viagem:</span>
            </h2>
            <% if modal <> true then %>
            <div class="text-right">
                <p id="cambio-label">Câmbio: R$ <%=formatnumber(cambio,2)%></p>
            </div>
            <% end if %>
            <div class="row">
                <div class="col-md-6">
                    <select name="categoria" id="categoria" class="form-select campo-customizado mb-3 verificaReceptivo" required>
                    <option value="">Categoria</option>
                        <%                            
                            while not rsCategoria.eof                   
                                if cint(categoria_id) = rsCategoria("id") then
                                    categoria_nome = rsCategoria("nome")
                                    Response.Write "<option value="&rsCategoria("id")&" selected>"&rsCategoria("nome")&"</option>"
                                else                                        
                                    if plano_id = "" then Response.Write "<option value="&rsCategoria("id")&">"&rsCategoria("nome")&"</option>"
                                end if                                                                                                       
                            rsCategoria.movenext
                            wend                                                                  
                        %>
                    </select>
                </div>
                <div class="form-group col-md-4" hidden="hidden" disabled="disabled" id="planoReceptivo">
                    <select class="form-select campo-customizado mb-3" id="planoPaises" name="planoPaises">
                        <option disabled selected value="">PAIS DE ORIGEM</option>
                        <%
                            while not rsPaises.eof
                                Response.Write "<option value="&rsPaises("cod")&">"&rsPaises("nome")&"</option>"
                            rsPaises.movenext
                            wend
                        %>
                    </select>
                </div>                  
                <div class="col-md-6">
                    <select name="destino" id="destino" class="form-select campo-customizado mb-3" required>
                    <option value="">Destino da viagem</option>
                        <%  
                            while not rsDestino.eof                   
                                if cint(destino_id) = rsDestino("destinoId") then
                                    destino_nome = rsDestino("destinoNome")
                                    Response.Write "<option value="&rsDestino("destinoId")&" selected>"&rsDestino("destinoNome")&"</option>"
                                else                                        
                                    Response.Write "<option value="&rsDestino("destinoId")&">"&rsDestino("destinoNome")&"</option>"
                                end if                                                                                                       
                            rsDestino.movenext
                            wend 
                        %>
                    </select>
                </div>
            </div>
            <div class="row">
            <div class="col-md-3">
                <input type="text"  class="form-control campo-customizado calendar upCovid mb-3" id="inicioViagem" name="inicioViagem" data-provide="datepicker" required="required" readonly="readonly" value="<%=data_inicio%>" placeholder="Embarque">
            </div>
            <div class="col-md-3">
                <div class="input-container">
                    <div class="recomendado" id="days">DIAS: <%=dias%></div>
                    <input type="text" class="form-control campo-customizado calendar mb-3" id="fimViagem" name="fimViagem" data-provide="datepicker" required="required" readonly="readonly" value="<%=data_fim%>" placeholder="Desembarque">
                </div>
            </div>
            <div class="form-group col-md-3 passageiros">
                <div class="recomendado" id="days">0 A 64 ANOS</div>
                <select name="idadeMenor" id="idadeMenor" class="form-select campo-customizado mb-3" required>
                <option value="">nº de passageiros</option>
                    <%
                        For i = 0 To 10
                            if i = nPax_novo then
                                Response.write "<option value="&i&" selected>"&i&" passageiros</option>"
                            else
                                Response.write "<option value="&i&">"&i&" passageiros</option>"
                            end if
                        Next                                
                    %>
                </select>
            </div>
            <div class="col-md-3">
                <div class="recomendado" id="days">65 A 85 ANOS</div>
                <select name="idadeMaior" id="idadeMaior" class="form-select campo-customizado mb-3" required>
                <option value="">nº de passageiros</option>
                    <%
                        For i = 0 To 10
                            if i = nPax_idoso then
                                Response.write "<option value="&i&" selected>"&i&" passageiros</option>"
                            else
                                Response.write "<option value="&i&">"&i&" passageiros</option>"
                            end if
                        Next
                    %>
                </select>
            </div>
   
            <div class="col-md-4">
                <input type="text" class="form-control campo-customizado mb-3" name="nomeCli" placeholder="Nome">
            </div>
            <div class="col-md-4">
                <input type="text" class="form-control campo-customizado mb-3" name="email" placeholder="E-mail">
            </div>

            <div class="col-md-4">
                <input type="text" class="form-control campo-customizado mb-3 telefone" name="fone" placeholder="Telefone com DDD">
            </div>
            </div>
            <div class="form-group col-12 radio">
                <label for="incluir">
                    Incluir a <span style="font-weight: 600; font-style: italic;">cobertura complementar para Covid-19</span> no
                    plano?
                </label>
                <div class="custom-control custom-switch col-md-2 pt-1" style="display:none;">
                    <input type="checkbox" class="custom-control-input upCancel upCovid verificaFamiliar" id="familiar" name="familiar" <% 
                            if familiar = 0 and familiar <> "" and plano_id <> "" then 
                                Response.Write("disabled") 
                            else 
                                if familiar = 1 then 
                                    Response.Write("checked") 
                                end if 
                            end if 
                        %>>
                    <label class="custom-control-label" for="familiar"><b>PLANO FRIENDS</b></label>
                </div>
                <%if not rsCovid.eof then%>
                            <input type="checkbox" class="custom-control-input upCovid" id="upgradeCovid"
                                name="upgradeCovid"
                                <%if upgrade_covid <> "" and upgrade_covid <> 0 then Response.Write("checked")%>>
                            <label class="custom-control-label" for="upgradeCovid"><b>COVID-19</b></label>
				<%end if%>
                <br>
                <label>Existem restrições de coberturas para países listados pelo <a
                    href="https://seguroviagemnext.com.br/Products/sobre-next.asp#OFAC"><strong>OFAC</strong></a></label>

                
                <div class="form-group col-md-3">
                    <select class="custom-select campo-customizado" id="planoCovid" name="planoCovid"
                        hidden="hidden" disabled="disabled">                                                               
                    </select>
                </div>  
            </div>
            <br>
            <div class="row">
                <div class="col-md-6 desconto">
                    <p>
                    *Desconto de até 70% <br>
                    Até 10x s/ juros no cartão
                    </p>
                </div>
                <input type="text" hidden="hidden" id="address" name="address" value="<%=address%>" />
                <% if plano_id <> "" then %>
                <input type="text" hidden="hidden" name="planoId" value="<%=plano_id%>" />
                <% end if %>
                <div class="col-md-6 text-end"><button type="submit" class="cta" >REALIZAR MINHA COTAÇÃO</a></div>
            </div>
        </div>
        </form>
        <script>
            function verificaCotacao() {
                // Sua lógica de verificação
                var boll = true;
                var message = "";
              
                if (document.getElementById('inicioViagem').value == '') {
                    // mostra um modal
                    $("#modalEmbarque").modal();
                    return false;
                }
                if(document.getElementById('fimViagem').value == ''){
                    $("#modalDesembarque").modal();
                    return false; // O formulário não será enviado
                }
                var idadeJoven = parseInt($("#idadeMenor").val());
                var idadeIdoso = parseInt($("#idadeMaior").val());

                if(isNaN(idadeJoven) ){
                    $("#modalPassageiroJ").modal();
                    return false;
                }if(isNaN(idadeIdoso) ){
                    $("#modalPassageiroI").modal();
                    return false;
                }
             
            }
        </script>
    </div>
    <!-- FIMMMMMMMM -->

    