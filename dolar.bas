Attribute VB_Name = "Módulo1"
Sub exemplo()

    Dim requisicao As Object
    Dim html As String
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    requisicao.Open "GET", "https://www.microsoft.com/pt-br/", False
    requisicao.send
    html = requisicao.ResponseText
    Debug.Print html
    
    Set requisitos = Nothing
    html = vbNullString
    

End Sub
