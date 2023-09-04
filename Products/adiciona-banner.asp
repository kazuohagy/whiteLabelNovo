<%
    response.cookies("agt")("id") = request.QueryString("id")
%>
<!DOCTYPE html>
<html>
<head>
    <title>Upload de Imagem</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <div class="card p-4">
            <form action="upload.asp" method="post" enctype="multipart/form-data">
                <div class="text-center mb-3">
                    <img src="https://seguroviagemnext.com.br/products/uploads/<%=request.QueryString("id")%>.jpg" class="img-fluid rounded" style="max-width: 200px; max-height: 200px;" alt="Imagem">
                </div>
                <div class="mb-3">
                    <a>Imagem (1024 x 649 px) *</a>
                    <input type="file" class="form-control" name="imagem" accept="image/*">
                </div>
                <div class="text-center">
                    <button type="submit" class="btn btn-primary">Enviar Imagem</button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>

