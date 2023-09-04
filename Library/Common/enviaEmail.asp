<%
	'Session.CodePage=65001

	logEnviaMail = "<B>DESCRICAO DO LOG DE EVENTOS</B><BR><BR><BR>"

	function enviaMail(emailRemetente,nomeRemetente,emailCopia,emailDestinatario,assunto,corpo,confirmacao)

	'enviar nomeRemetente e emailRemetente em branco 
	'caso queira que seja preenchido o padrão
	if emailRemetente = "" then
		emailRemetente = "no-reply@nextseguroviagem.com.br"
	end if
	if nomeRemetente = "" then
		
		nomeRemetente = "Next Seguro Viagem"
	end if
		
	emaildestino = emailDestinatario ' e-mail que vai receber as mensagens do formulario

	Set JMail = Server.CreateObject ("JMail.message")

	JMail.Logging = true
	JMail.silent = true	
	JMail.Priority = 3
	JMail.AddNativeHeader  "X-Mailer", "TAVOLA, Build 11.0.5510" 
	JMail.AddNativeHeader "X-MimeOLE", "Produced By FICOPOLA MimeOLE V6.00.2900.2869" 

if not isnumeric(confirmacao) then

	if INSTR(confirmacao,"|") then

	vet_confirmacao = SPLIT(confirmacao,"|")


For iA = 0 to ubound(vet_confirmacao)
	if vet_confirmacao(iA) <> "" then 	JMail.AddAttachment vet_confirmacao(iA)
NEXT

else

vet_confirmacao = SPLIT(confirmacao,",")
vet_confirmacao = SPLIT(confirmacao,"|")


	For iA = 0 to ubound(vet_confirmacao)
	if vet_confirmacao(iA) <> "" then 	JMail.AddAttachment vet_confirmacao(iA)
		NEXT
	end if
end if
	
	JMail.logging=True
	JMail.From = emailRemetente
	JMail.FromName = nomeRemetente
	JMail.AddRecipient emailDestinatario
	JMail.ISOEncodeHeaders = false 
	JMail.ReplyTo = emailRemetente
	JMail.Subject = assunto
	JMail.HTMLBody = corpo
	JMail.MailServerUserName = "nextseguroviagem" ' conta de e-mail utilizada para enviar
	JMail.MailServerPassWord = "oBqWNSuL4048" ' senha da conta de e-mail
	JMail.Send("smtplw.com.br:587") ' Informacoes so seu servidor SMTP	
	'talvez seja necessário alterar a porta
	
	logEnviaMail = logEnviaMail & JMail.Log  & "<BR><BR>"			
end function



Function FncBinCheckMail(StrMail)
	' FunÂ�Â�o que verifica validaÂ�Â�o de preenchimento de E-Mail.
	
	' Se hÂ� espaÂ�o vazio, entÂ�o...
	If InStr(1, StrMail, " ") > 0 Then
	FncBinCheckMail = False
	Exit Function
	End If
	
	' Verifica tamanho da String, pois o menor endereÂ�o vÂ�lido Â� x@x.xx.
	If Len(FncStrSpace(StrMail)) < 6 Then
	FncBinCheckMail = False
	Exit Function
	End If
	' Verifica se se hÂ� um "@" no endereÂ�o.
	If InStr(FncStrSpace(StrMail), "@") = 0 Then
	FncBinCheckMail = False
	Exit Function
	End If
	' Verifica se hÂ� um "." no endereÂ�o.
	If InStr(FncStrSpace(StrMail), ".") = 0 Then
	FncBinCheckMail = False
	Exit Function
	End If
	' Verifica se hÂ� a quantidade mÂ�nima de caracteres Â� igual ou maior que 3.
	If Len(FncStrSpace(StrMail)) - InStrRev(FncStrSpace(StrMail), ".") > 3 Then
	FncBinCheckMail = False
	Exit Function
	End If
	
	' Verifica se hÂ� "_" apÂ�s o "@".
	If InStr(FncStrSpace(StrMail), "_") <> 0 And InStrRev(StrMail, "_") > InStrRev(FncStrSpace(StrMail), "@") Then
	FncBinCheckMail = False
	Exit Function
	Else
	Dim IntCounter
	Dim IntF
	IntCounter = 0
	For IntF = 1 To Len(FncStrSpace(StrMail))
	If Mid(StrMail, IntF, 1) = "@" Then
	IntCounter = IntCounter + 1
	End If
	Next
	If IntCounter > 1 Then
	FncBinCheckMail = True
	End If
	' Valida cada caracter do endereÂ�o.
	IntF = 0
	For IntF = 1 To Len(FncStrSpace(StrMail))
	If IsNumeric(Mid(FncStrSpace(StrMail), IntF, 1)) = False And _
	(LCase(Mid(FncStrSpace(StrMail), IntF, 1)) < "a" Or _
	LCase(Mid(FncStrSpace(StrMail), IntF, 1)) > "z") And _
	Mid(FncStrSpace(StrMail), IntF, 1) <> "_" And _
	Mid(FncStrSpace(StrMail), IntF, 1) <> "." And _
	Mid(FncStrSpace(StrMail), IntF, 1) <> "-" Then
	FncBinCheckMail = True
	End If
	Next
	End If
End Function

Function FncStrSpace(StrAddress)
	FncStrSpace = Trim(LTrim(RTrim(StrAddress)))
End Function
%>