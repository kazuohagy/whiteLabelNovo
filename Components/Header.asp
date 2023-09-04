<!-- Início navegador -->
<nav class="navbar navbar-expand-lg bg-white">
  <div class="container">
    <a class="navbar-brand" href="https://seguroviagemnext.com.br/<%=request.Cookies("wlabel")("siteNome")%>"><img id="logo-marca" src="https://www.nextseguro.com.br/NextSeguroViagem/old/img/agtaut/<%=request.Cookies("wlabel")("revId")%>.jpg" style="width: 90px;"></a>
    <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarSupportedContent"
      aria-controls="navbarSupportedContent" aria-expa l="Toggle navigation">
      <span class="navbar-toggler-icon"></span>
    </button>
    <div class="collapse navbar-collapse" id="navbarSupportedContent">
      <ul class="navbar-nav ms-auto mb-2 mb-lg-0">
        <li class="nav-item">
          <a class="nav-link fs-5" aria-current="page" href="https://seguroviagemnext.com.br/<%=request.Cookies("wlabel")("siteNome")%>">Home</a>
        </li>
        <li class="nav-item">
          <a class="nav-link fs-5" href="https://seguroviagemnext.com.br/Products/sobre-next.asp">Sobre a NEXT</a>
        </li>
        <li class="nav-item">
          <a class="nav-link fs-5" href="https://seguroviagemnext.com.br/Products/diferenciais.asp">Diferenciais</a>
        </li>
      </ul>
    </div>
  </div>
</nav>
<!-- Fim navegador -->