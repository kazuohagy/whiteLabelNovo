
<!--#include  file="../Library/Common/micMainCon.asp" -->
<HTML> 
	<head>
		<script src="https://cdn.lordicon.com/bhenfmcm.js"></script>
		 <!--#include file="../Components/HTML_Head2.asp" --> 
	</head>
<BODY> 
<%
link = Server.MapPath("teste.asp")
'response.write link
'response.end
Dim objUpload, NewName, ObjRs, File
Server.ScriptTimeout = 3000000

	Set objUpload = server.CreateObject("Dundas.Upload.2") 

if request.cookies("agt")("id") <> "" then

objConn.Execute("UPDATE goAffinity_afiliado SET bannerNome='"&request.cookies("agt")("id")&"' WHERE clienteId ='"&request.cookies("agt")("id")&"'")
objConn.close
Set objFS = Server.CreateObject("Scripting.FileSystemObject")
If not objFS.FolderExists("C:\Inetpub\vhosts\affinityseguro.com.br\seguroviagemnext.com.br\products\uploads\") Then
' criar a pasta
objFS.CreateFolder("C:\Inetpub\vhosts\affinityseguro.com.br\seguroviagemnext.com.br\products\uploads\")
Set objFS = Nothing
End if
end if

	NewName = request.cookies("MR")("foto")
	
	if err.number <> 0 then 
	Response.write  err.description 
	end if 

	'estipula o tamanho máximo do arquivo 
	objUpload.MaxFileSize = 1048576 
	
	'formatando o nome do arquivo 
	objUpload.UseUniqueNames = false 
	
	'informa o path onde os arquivos serão salvos 
	'obs: o diretório deve ter permissão de escrita 
	'objUpload.Save "E:\vhosts\neisazulian.com\httpdocs\tavola\artigos\filesNZ" 
	objUpload.SaveToMemory

	objUpload.Files(0).SaveAs "C:\Inetpub\vhosts\affinityseguro.com.br\seguroviagemnext.com.br\Products\uploads\" & request.cookies("agt")("id") & ".jpg"
	
%> 

<div class="card text-center">
      <div class="card-body">
        <h1 class="card-title">Imagem Adicionada com Sucesso!</h1>
		<lord-icon
    src="https://cdn.lordicon.com/tyvtvbcy.json"
    trigger="hover"
    colors="primary:#121331"
    style="width:250px;height:250px">
</lord-icon>
        <p class="card-text">Ótimo trabalho! O Banner foi adicionada com sucesso.</p>
        <div class="animated-ok"></div>
      
       
      </div>
    </div>
<br>
<%
Dim Jpeg, Path, divChave

'Criando o full: 
Set Jpeg = Server.CreateObject("Persits.Jpeg") 
'Caminho da Imagem 
Path = "C:\Inetpub\vhosts\affinityseguro.com.br\seguroviagemnext.com.br\Products\uploads\" & request.cookies("agt")("id") & ".jpg"
'Busca a Imagem 
Jpeg.Open Path 
'Especifica o tamanho da imagem

'primeiramente define uma largura fixa
'divChave = Jpeg.OriginalWidth / 160
'Jpeg.Height = Jpeg.OriginalHeight / divChave
'Jpeg.Width = 160

'if cdbl(Jpeg.Height) > 115 then
''caso a altura passe de 115 px, define uma altura fixa.
'divChave = Jpeg.OriginalHeight / 105
'Jpeg.Width = Jpeg.OriginalWidth / divChave
'Jpeg.Height = 105
'end if

'Esse método é opcional, usado para melhorar o visual da imagem 
Jpeg.Sharpen 1, 150 
'Cria um thumbnail e o grava no caminho abaixo 
Jpeg.Save "C:\Inetpub\vhosts\affinityseguro.com.br\seguroviagemnext.com.br\Products\uploads\" & request.cookies("agt")("id") & ".jpg"


'Para enviar o thumbnail para o browser do cliente utilize o método SendBinary: 
'Response.Write jpeg.SendBinary 

'objconn.execute("UPDATE cadCliente SET logo='1' WHERE id='"&request.cookies("FCNET")("idAge")&"'")

Set Jpeg = Nothing
%>
<script>
window.opener.document.location.reload();
</script>
</BODY>
</HTML>

