<!--#include file="../../biblioteca/MainCon.asp" -->
<!--#include file="../../biblioteca/funcoes.asp" -->
<%

processoId = request.QueryString("processoId")

%>
<html>
<head>
<link rel="stylesheet" media="all" href="../css/normalize.css">
<link rel="stylesheet" media="all" href="../css/cake_core.css">
<link rel="stylesheet" media="all" href="../css/smoothness/jquery-ui-1.8.5.custom.css">
</head>

<body>
<form name="form1" method="post" action="processarPagamento.asp">
  <table width="100" border="0" align="center" cellpadding="2" cellspacing="0" style="width:420px; height:336px;">
    <tr>
      <td width="27" rowspan="3">&nbsp;</td>
      <td width="346" height="81"><h3>Intermac Assist&ecirc;ncia</h3></td>
      <td width="35" rowspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td height="209" valign="top"><p><strong>HOMOLOGA&Ccedil;&Acirc;O - COMPRA N&Atilde;O V&Aacute;LIDA</strong></p>
        <p><strong>Compra de seguro viagem:<br>
          De: 01/04/2014 At&eacute; 02/04/2014<br>
          Vigencia: 1 dia<br>
          Valor
          : R$ 1,00</strong><br>
          <br>
          1 Parcela
        </p>
        <table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="#FCFCFC" >
          <tr valign="top" >
            <td colspan="8"><b><font size="1"> Selecione a forma de Pagamento</font></b></td>
          </tr>
          <tr valign="top" >
            <td colspan="8"><h3>Cr&eacute;dito: </h3></td>
          </tr>
          <tr valign="top">
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche"  id="TRVisa" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif"><label for="radio2">
                <input name="pagamento" type="radio" id="radio2" value="VI">
                <br>
                <img src="../../img/Cartoes/1347651304_visa-curved.png" width="102" height="64" onClick="document.getElementById('radio2').checked=true;"/> </label></td>
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche" id="TRMaster" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif"><label for="radio3">
              <input type="radio" name="pagamento" id="radio3" value="MA"  >
              <br>
              <img src="../../img/Cartoes/1347651307_mastercard-curved.png" width="102" height="64"  onClick="document.getElementById('radio3').checked=true;"/> </label></td>
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche" id="TRMaster" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif"><label for="radio4">
              <input type="radio" name="pagamento" id="radio4" value="AM"  >
              <br>
              <img src="../../img/Cartoes/1347651301_american-express-curved.png" width="102" height="64"  onClick="document.getElementById('radio4').checked=true;"/> </label></td>
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche" id="TRMaster" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif"><label for="radio5">
              <input type="radio" name="pagamento" id="radio5" value="DI"  >
              <br>
              <img src="../../img/Cartoes/1347651304_dinners.png" width="102" height="64"  onClick="document.getElementById('radio5').checked=true;"/></label></td>
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche" id="TRMaster" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif"><label for="radio6">
              <input type="radio" name="pagamento" id="radio6" value="DS"  >
              <br>
              <img src="../../img/Cartoes/1347651301_discover.png" width="102" height="64"  onClick="document.getElementById('radio6').checked=true;"/></label></td>
            <td align="center" valign="middle" bgcolor="#F0EEEE"><label for="radio7">
            <input type="radio" name="pagamento" id="radio7" value="JC"  >
              <br>
            <img src="../../img/Cartoes/1347651301_jcb.png" width="102" height="64"  onClick="document.getElementById('radio6').checked=true;"/>
            </label>
            </td>
            <td align="center" valign="middle" bgcolor="#F0EEEE">
            <label for="radio8">
            <input type="radio" name="pagamento" id="radio8" value="AU"  >
              <br>
            <img src="../../img/Cartoes/1347651301_aura.png" width="102" height="64"  onClick="document.getElementById('radio6').checked=true;"/>
            </label></td>
            <td align="center" valign="middle" bgcolor="#F0EEEE">
            <label for="radio9">
            <input type="radio" name="pagamento" id="radio9" value="EL"  >
              <br>
            <img src="../../img/Cartoes/1347651304_elo-curved.png" width="102" height="64"  onClick="document.getElementById('radio6').checked=true;"/>
            </label></td>
          </tr>
          <tr valign="top" >
            <td colspan="8"><h3>D&eacute;bito: </h3></td>
          </tr>
          <tr valign="top">
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche"  id="TRVisa2" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif">
            <label for="radio10">
              <input name="pagamento" type="radio" id="radio10" value="DV">
              <br>
              <img src="../../img/Cartoes/1347651304_visa-curved.png" width="102" height="64" onClick="document.getElementById('radio2').checked=true;"/></label></td>
            <td align="center" bgcolor="#F0EEEE" class="tdPreenche" id="TRMaster2" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif">&nbsp;</td>
            <td colspan="6" align="center" bgcolor="#F0EEEE" class="tdPreenche" id="TRMaster2" style="border-right:#CCCCCC 1px solid; border-bottom:#CCCCCC 1px solid; font-size:12px; font-family:'Times New Roman', Times, serif">&nbsp;</td>
          </tr>
          <tr valign="top">
            <td colspan="8" align="center" valign="top"><input name="button" type="submit" class="botaoAzul" id="button" value="Continuar"></td>
          </tr>
        </table>
        <p><br>
          <br>
        </p></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>

