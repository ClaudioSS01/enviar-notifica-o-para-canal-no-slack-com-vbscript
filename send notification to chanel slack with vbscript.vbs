Function SendMessageToSlack(message)
    'https://humanitargroup.slack.com/apps/A0F7XDUAZ-webhooks-de-entrada?tab=more_info
    'o linkl acima é para vc adicionar o webhook ao seu canal do slack
    Dim slackURL  : slackURL = "url do webhook"
    Dim slackChannel : slackChannel = "nome do canal"
    Dim slackUsername : slackUsername = "webhookbot"
    Dim slackIcon : slackIcon = ":ghost:"
    
    ' Crie um objeto de solicitação HTTP e configure o cabeçalho
    Dim httpRequest : Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "POST", slackURL, False
    httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Monte o payload da mensagem do Slack
    Dim payload : payload = "payload={""channel"": """ & slackChannel & """, ""username"": """ & slackUsername & """, ""text"": """ & message & """, ""icon_emoji"": """ & slackIcon & """}"
    
    ' Envie a mensagem para o Slack
    httpRequest.send payload
    
    ' Verifique se a mensagem foi enviada com sucesso
    If httpRequest.Status = 200 Then
        MsgBox "Mensagem enviada com sucesso para o canal " & slackChannel
    Else
        MsgBox "Falha ao enviar mensagem: " & httpRequest.responseText
    End If
    
    ' Libere o objeto de solicitação HTTP
    Set httpRequest = Nothing
End Function

' Exemplo de uso:
SendMessageToSlack "Funcionando com sucesso!"
