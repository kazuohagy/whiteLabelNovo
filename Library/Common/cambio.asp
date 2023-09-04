<%
    Session.CodePage=65001

Dim cambioTarifaRS, cambio
set cambioTarifaRS = objConn.execute("select top 1 * from cadCambio where data <= GETDATE() order by data desc")

    if not cambioTarifaRS.eof then
        cambio = cambioTarifaRS("usdMic")
    end if

%>