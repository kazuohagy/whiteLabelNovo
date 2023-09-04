<!--#include file="../../conexoes/master.asp"-->
<!--#include file="../conexoes/tavola.asp" -->
<!--#include file="../conexoes/ficopola.asp" -->
<!--#include file="../funcoesInc/funcoes.asp" -->
<html>
<head>
<title></title>
<LINK REL=stylesheet HREF="../css/main.css" TYPE="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="../includes/barra_cadastro.asp"--><br>
<% response.Buffer=true
if request.QueryString("confirmacao")="" then%>
<table width="775" border="0" align="center" cellpadding="4" cellspacing="0">
  <tr> 
    <td valign="middle"><b><font size="4">
      Importa&ccedil;&atilde;o de Clientes - Site ficopola.net
			
    </font></b><br><br><div align="center">
    <input type="button" class="botaoVerde" value="Iniciar importação >>" onClick="document.location='importar.asp?confirmacao=1'"></div></td>
  </tr>
</table>
<%else

'sqlStr="from cadCliente where id = '2022'"
sqlStr="from cadCliente where tt='0' or tt is null"

set totalRS = objconnFI.execute("select count(id) "&sqlStr)

set objRS  =  objconnFI.execute("select * "&sqlStr&" order by id")


%>



<table width="775" border="0" align="center" cellpadding="4" cellspacing="0">
  <tr> 
    <td width="614" valign="middle"><b><font size="4">
      Importa&ccedil;&atilde;o de Clientes - Site ficopola.net
			
    </font></b></td>
    <td width="145" valign="bottom"><div id="divLoading" align="left" style=' padding-left:6px; padding-top:3px; font-size:10px; background-color:#FFECEC; width: 135; height: 20px; color:#990000; font-weight:bold; font-family:Verdana; font-size:10px; '>Carregando... (1%)</div></td>
  </tr>
  <tr>
    <td colspan="2" valign="middle"><hr></td>
  </tr>
</table>

<table width="775" border="0" cellpadding="2" align="center" cellspacing="0">
  <%i=0
 if objrs.eof then %>
  <tr>
    <td colspan=10><font size=2><b>Nenhum novo cliente foi encontrado.</b></font></td>
  </tr>
  <%	else 
  %>
    

  <tr bgcolor="#6893C6" style="color:#FFFFFF"> 
    <td width="46"><font color="#FFFFFF"><b>Cod</b></font></td>
    
    <td width="66" bgcolor="#6893C6"><font color="#FFFFFF"><b>Cliente</b></font></td>
    <td width="90"><font color="#FFFFFF"><b>Cidade</b></font></td>
    <td width="138"><b><font color="#FFFFFF">Telefone</font></b></td>
    <td><b><font color="#FFFFFF">Endere&ccedil;o</font></b></td>
  </tr>
  
  <% 

		while not objrs.eof
				
if not (isnull(objRs("dataCadastro")) or isEmpty(objRs("dataCadastro"))) then dataCadastro=data(DAY(objRs("dataCadastro"))&"/"&MONTH(objRs("dataCadastro")) &"/"& YEAR(objRs("dataCadastro")),3,0) else dataCadastro=""

fantasia=replace(objRs("fantasia"),"'","")
razao=replace(objRs("razao"),"'","")
if not isnull(objRs("endereco")) then endereco=replace(objRs("endereco"),"'","")
if not isnull(objRs("bairro")) then bairro=replace(objRs("bairro"),"'","")
if not isnull(objRs("cidade")) then cidade=replace(objRs("cidade"),"'","")
		
'---insere a agencia no BD econumus


objconn.execute("insert into cadCliente (id,fantasia,razao,tipo,endereco,contatoPrincipal,email,bairro,complemento,cidade,uf,cep,pais,fone,fax,cnpj,insc,site,bancoId,agencia,conta,responsavelCadastro,origemCadastro,obs,cobFantasia,cobRazao,cobNome,cobEndereco,cobBairro,cobComplemento,cobCidade,cobUF,cobCEP,cobCNPJ,cobFone,cobFAX,cobcontato,favorecido,ativo,nf,dataCadastro,cobEmail) values ('"&objRs("id")&"','"&fantasia&"','"&razao&"','"&objRs("tipo")&"','"&endereco&"','"&objRs("contatoPrincipal")&"','"&objRs("email")&"','"&bairro&"','"&objRs("complemento")&"','"&cidade&"','"&objRs("uf")&"','"&objRs("cep")&"','"&objRs("pais")&"','"&objRs("fone")&"','"&objRs("fax")&"','"&objRs("cnpj")&"','"&objRs("insc")&"','"&objRs("site")&"','"&objRs("bancoId")&"','"&objRs("agencia")&"','"&objRs("conta")&"','"&objRs("responsavelCadastro")&"','"&objRs("origemCadastro")&"','"&objRs("obs")&"','"&objRs("cobFantasia")&"','"&objRs("cobRazao")&"','"&objRs("cobNome")&"','"&objRs("cobEndereco")&"','"&objRs("cobBairro")&"','"&objRs("cobComplemento")&"','"&objRs("cobCidade")&"','"&objRs("cobUF")&"','"&objRs("cobCEP")&"','"&objRs("cobCNPJ")&"','"&objRs("cobFone")&"','"&objRs("cobFAX")&"','"&objRs("cobcontato")&"','"&objRs("favorecido")&"','"&objRs("ativo")&"','"&objRs("nf")&"','"&dataCadastro&"','"&objRs("email")&"')")

objconnFI.execute("UPDATE cadCliente set tt='1' WHERE id="&objRs("id"))

'------------------------
			i = i + 1

			cor = cor MOD 2
			if cor = 0 then
				BG="#FFFFFF"
				bg2="#E0F4E3"
			else
				BG="#C8DADB"
				bg2="#D5E9D0"
			end if
			
			cor = cor + 1 
			porcentagem = formatNumber(i/totalRs(0)*100,0)%>

<script>document.getElementById('divLoading').innerHTML='Carregando... (<%=porcentagem%>%)'</script>

  <tr bgcolor="<%=bg%>">
    <td height="22"><%=objrs("id")%></td>
    
    <td><a href="administra.asp?id=<%=objrs("id")%>"><%=objrs("razao")%></a></td>
    <td><%=objrs("cidade")%></td>
    <td><%=objrs("fone")%></td>
    <td><%=objrs("endereco")%></td>
  </tr>
  <%response.Flush()
  			objrs.movenext
		wend
	end if %>
<script>document.getElementById('divLoading').style.display='none'</script>
</table>
<table width="775" border="0" align="center" cellpadding="4" cellspacing="1">
  <tr>
    <td align="right" bgcolor="#6893C6" width="45%"><font size="2" color="#FFFFFF"><b>Quantidade de ag&ecirc;ncias importados:</b></font></td> 
    <td valign="top"><font size="2"><b><%=i%></b></font></td>
  </tr>
</table>


<table width="775" border="0" align="center" cellpadding="4" cellspacing="0">
  <tr>
    <td width="614" valign="middle"><b><font size="4"> Importa&ccedil;&atilde;o de Contatos - Site ficopola.net </font></b></td>
    <td width="145" valign="bottom"><div id="divLoading2" align="left" style=' padding-left:6px; padding-top:3px; font-size:10px; background-color:#FFECEC; width: 135; height: 20px; color:#990000; font-weight:bold; font-family:Verdana; font-size:10px; '>Carregando... (1%) <font color=black>cod. 00000</font> </div></td>
  </tr>
  <tr>
    <td colspan="2" valign="middle"><hr></td>
  </tr>
</table>
<table width="775" border="0" cellpadding="2" align="center" cellspacing="0">
  <%i=0

sqlStr="from clienteContato where tt='0'"

set totalRS = objconnFI.execute("select count(id) "&sqlStr)
set objRS  =  objconnFI.execute("select * "&sqlStr&" order by id")

 if objrs.eof then %>
  <tr>
    <td colspan=8><font size=2><b>Nenhum novo emissor foi encontrado.</b></font></td>
  </tr>
  <%	else 
  %>
  <tr bgcolor="#6893C6" style="color:#FFFFFF">
    <td width="78"><font color="#FFFFFF"><b>Cod</b></font></td>
    <td width="285" bgcolor="#6893C6"><font color="#FFFFFF"><b>Nome</b></font></td>
    <td><b><font color="#FFFFFF">Ag&ecirc;ncia ID </font></b></td>
  </tr>
  <% i=0

		while not objrs.eof
		
			i = i + 1

			cor = cor MOD 2
			if cor = 0 then
				BG="#FFFFFF"
				bg2="#E0F4E3"
			else
				BG="#C8DADB"
				bg2="#D5E9D0"
			end if
			
			cor = cor + 1 
			porcentagem = formatNumber(i/totalRs(0)*100,0)%>
  <script>document.getElementById('divLoading2').innerHTML='Carregando... (<%=porcentagem%>%) <font color=black>cod. <%=objrs("id")%></font>'</script>
  <%
nome=objRs("nome")
email=objRs("email")
cargo=objRs("cargo")

login=objRs("login")

if not isnull(nome) then nome=replace(nome,"'","")
if not isnull(email) then email=replace(email,"'","")
if not isnull(cargo) then cargo=replace(cargo,"'","")
if not isnull(bancoNome) then bancoNome=replace(bancoNome,"'","")
if not isnull(login) then login=replace(login,"'","")

ultimoAcesso = objRs("ultimoAcesso")

if not isnull(ultimoAcesso) then ultimoAcesso = data(objRs("ultimoAcesso"),3,0)

		
'---insere o contato no BD Harmonica

objconn.execute("insert into clienteContato (id,nome,aniversario,email,cargo,idCliente,login,senha,nivel,ultimoAcesso,acessos,ativo) values ('"&objRs("id")&"','"&nome&"','"&objRs("aniversario")&"','"&email&"','"&cargo&"','"&objRs("idCliente")&"','"&objRs("login")&"','"&objRs("senha")&"','"&objRs("nivel")&"','"&ultimoAcesso&"','"&objRs("acessos")&"','"&objRs("ativo")&"') ")

objconnFI.execute("UPDATE clienteContato set tt='1' WHERE id="&objRs("id"))

'------------------------
  %>
  <tr bgcolor="<%=bg%>">
    <td height="22"><%=objrs("id")%></td>
    <td><%=objrs("nome")%></td>
    <td><%=objrs("idCliente")%></td>
  </tr>
  <%response.Flush()
  			objrs.movenext
		wend
	end if %>
  <script>document.getElementById('divLoading2').style.display='none'</script>
</table>
<table width="775" border="0" align="center" cellpadding="4" cellspacing="1">
  <tr>
    <td align="right" bgcolor="#6893C6" width="45%"><font size="2" color="#FFFFFF"><b>Quantidade de emissores importados:</b></font></td>
    <td valign="top"><font size="2"><b><%=i%></b></font></td>
  </tr>
</table>
<%end if%> 
</body>
</html>
