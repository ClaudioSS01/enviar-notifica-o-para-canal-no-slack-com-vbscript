const fs = require('fs');
const {
  exec
} = require('child_process');






function cmd(comandoDeCmd = "tree") {
  exec(comandoDeCmd, (err, stdout, stderr) => {
    if (err) {
      console.error(err);
      return;
    }
    console.log(stdout);
  });
}




function log(texto) {
  console.log(texto);
}



// Obtém a data/hora atual
let data = new Date();

// Guarda cada pedaço em uma variável
let dia     = data.getDate();           // 1-31
let mes     = (data.getMonth() + 1);          // 0-11 (zero=janeiro)
let ano4    = data.getFullYear();       // 4 dígitos
let hora    = data.getHours();          // 0-23
let minuto     = data.getMinutes();        // 0-59
let seg     = data.getSeconds();        // 0-59


// Formata a data e a hora (note o mês + 1)
const hoje = dia + '/' + (mes+1) + '/' + ano4;
let str_hora = hora + ':' + minuto + ':' + seg;


// Mostra o resultado
console.log('Hoje é ' + hoje + ' às ' + str_hora);



(async () => {
  log('Iniciando script para enviar o relatorio baixado por email');
function enviarNotificacaoNoSlack(mensagemParaNotificarNoSlac) {
let comando = `
Function SendMessageToSlack(message)
    Dim slackURL  : slackURL = "url pega no webhooks"
    Dim slackChannel : slackChannel = "nome do seu canal no slack com #"
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
        echo "Mensagem enviada com sucesso para o canal " & slackChannel
    Else
        echo "Falha ao enviar mensagem: " & httpRequest.responseText
    End If
    
    ' Libere o objeto de solicitação HTTP
    Set httpRequest = Nothing
End Function

' Exemplo de uso:
SendMessageToSlack "${mensagemParaNotificarNoSlac}"

`;

  //guardando a copia do historico
  fs.writeFile('sendNotificacaoParaSlack.txt', comando, err => {
    if (err) throw err;
  });

  //versao que vai ser executada
  fs.writeFile('tmp.vbs', comando, err => {
    if (err) throw err;
  });

  //`echo CreateObject("WScript.Shell").SendKeys "%{UP}" > tmp.vbs && cscript tmp.vbs && del tmp.vbs`
  let comandoParaEnviaroEmail = `cscript tmp.vbs && del tmp.vbs`;
  cmd(comandoParaEnviaroEmail);
  log('Fim da execução do script de enviar relatorio por email');
}

const parametros = "mensagem que você deseja enviar";

let mensagem = `Mensagem gerada em ${dia}/${mes}/${ano4} as ${hora}:${minuto}. Mensagem: ${parametros}  `

enviarNotificacaoNoSlack(mensagem)
})();
