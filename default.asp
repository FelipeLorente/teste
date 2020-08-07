aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa
<%
'set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP.3.0")
' 
'xmlhttp.Open "POST", "https://www.receitaws.com.br/v1/cnpj/62281795000100", False
'
'xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'xmlhttp.send DataToSend 
'
'Set xmlhttp = nothing




Response.Buffer = True

Usuario = Request.Form("usuario")
Senha = Request.Form("senha")
chassi = request.Form("chassi")


Dim xmlhttp, envelopeXML
'envelopeXML = (Usuario&Senha&Chassi)
Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
xmlhttp.open "POST", "https://www.receitaws.com.br/v1/cnpj/62281795000100", false
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'Supondo que a variável envelopeXML já possua um conteúdo.
xmlhttp.send()

response.Write xmlhttp.status
response.End()

'Estado 4: requisição foi feita e completada sem falhas. Status 200: comunicação realizada com êxito junto ao webservice.
If xmlhttp.readystate = 4 And xmlhttp.status = 200 Then
      Dim recebeXML

      'Criado um DOM para poder receber o arquivo XML e navegar dentro dele.
      Set recebeXML = Server.CreateObject("Microsoft.XMLDOM")
      recebeXML.setProperty "ServerHTTPRequest", True
      recebeXML.async = false
      recebeXML.LoadXML(xmlhttp.responseXML.xml)

      If recebeXML.parseError.errorCode <> 0 Then
            Dim percorreXML

            'Carregando nó principal
            Set percorreXML = recebeXML.documentElement 

            'Teste - Exibindo conteúdo do XML
            response.write percorreXML.childNodes.item(0).childNodes.item(0).attributes(0).text
      End If
End If
Set xmlhttp = nothing
Set recebeXML = nothing
Set percorreXML = nothing







'Set oJSON = New aspJSON
'
''Load JSON string
'oJSON.loadJSON(jsonstring)
'
''Get single value
'Response.Write oJSON.data("firstName") & "<br>"
'
''Loop through collection
'For Each phonenr In oJSON.data("phoneNumber")
'    Set this = oJSON.data("phoneNumber").item(phonenr)
'    Response.Write _
'    this.item("type") & ": " & _
'    this.item("number") & "<br>"
'Next
'
''Update/Add value
'oJSON.data("firstName") = "James"
'
''Return json string
'Response.Write oJSON.JSONoutput()
%>

<script>
//var xhr = new XMLHttpRequest();
//xhr.open('GET', 'https://www.receitaws.com.br/v1/cnpj/62281795000100', true);
//
//alert(xhr.status);
//
//xhr.send(data);


var xmlhttp = new XMLHttpRequest();
xmlhttp.onreadystatechange = function() {

  if (this.readyState == 4 && this.status >= 200) {
    var myObj = JSON.parse(this.responseText);
    alert(this.responseText)
  }
};

xmlhttp.open("GET", "www.receitaws.com.br/v1/cnpj/62281795000100", true);
xmlhttp.send();

</script>
