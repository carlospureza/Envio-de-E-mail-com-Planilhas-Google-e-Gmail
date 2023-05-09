// Variável que recebe a instancia da planilha ativa.

var plan = SpreadsheetApp 
.getActiveSpreadsheet(); 

// variáveis que recebem as células do cabeçalho de ("A1 à F1").

var A1 = plan.getRange("A1");

var B1 = plan.getRange("B1");

var C1 = plan.getRange("C1");

var D1 = plan.getRange("D1");

var E1 = plan.getRange("E1");

var F1 = plan.getRange("F1");

var mes= ["",
"Janeiro",
"Fevereiro",
"Março",
"Abril",
"Maio",
"Junho",
"Julho",
"Agosto",
"Setembro",
"Outubro",
"Novembro",
"Dezembro"];

var atual= mes[Utilities.formatDate(new Date(), "GMT -03:00", "M")]
+Utilities.formatDate(new Date(), "GMT -03:00", "/YYYY");

var anterior = mes[Utilities.formatDate(new Date(), "GMT -03:00", "M") -1]+" e ";

var ancestral = mes[Utilities.formatDate(new Date(), "GMT -03:00", "M") -2]+", ";

function enviarEmail(){
  
var dados= plan.getActiveSheet().getDataRange().getValues();

  for(let l = 0; l < dados.length; l++){

    if(l > 0){
var matrícula = dados[l][0];

var nome = dados[l][1];

var status = dados[l][2];

var email = dados[l][3];

switch(status){
case "Suspenso (Financeiramente)": var template = HtmlService.createTemplateFromFile("template_1");

var mensalidades = anterior +atual;

var prazo = Utilities.formatDate( new Date( new Date( new Date().setDate(10)).getTime() +2592000000), "GMT -03:00", "10/MM/YYYY");

var assunto = "Suspensão do Quadro de Sócios";
break;

case "Eliminado": var template = HtmlService.createTemplateFromFile("template_2");

var mensalidades = ancestral +anterior +atual;;

var prazo = Utilities.formatDate( new Date( new Date().setDate( new Date().getDate() +10)), "GMT -03:00", "dd/MM/YYYY");

var assunto = "Eliminação do Quadro de Sócios";

break;

}


template.matricula = matrícula;

template.nome = nome;

template.mensalidades = mensalidades;

template.prazo = prazo;

var msg = template.evaluate().getContent();

if(dados[l][3]!=""){
MailApp.sendEmail({to: email,
subject: assunto,
htmlBody: msg});

}
else{
var template = HtmlService.createTemplateFromFile("template_3");

template.mat = dados[l][0];

template.soc = dados[l][1];

var msg = template.evaluate().getContent();

var op= plan.getEditors();
for(let i = 0 ; i < op.length; i++){
  var email = op[i];
}
MailApp.sendEmail({to: email,
subject: "Falha de operação.",
htmlBody: msg});


}
    }
  }
}