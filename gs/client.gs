var planilha;
var guiaLista;
var ultimalinha;
var statusExecution = "prod";

//Trigger Formulario
function setUpTrigger() {
  ScriptApp.newTrigger('main')
  .forForm('1G0DCbbsl9qmxG04F9EVk6gBFCDplgjjC4WSb6r0SDSs')
  .onFormSubmit()
  .create();
}

//---------------------------------------------------------------- 
//Constroi Lista Elementos
function buildElements(){
    planilha = SpreadsheetApp.getActiveSpreadsheet();
    guiaLista = planilha.getSheetByName("Respostas"); //Resposta
  
  ultimalinha = guiaLista.getLastRow();
  
  //Coluna dos Elementos na planilha
  var madeira = guiaLista.getRange(ultimalinha,7).getValue().toString();  
  var intMadeira = 0
  if(madeira.length > 0)
    var intMadeira = madeira.split(", ").length;

  var fogo = guiaLista.getRange(ultimalinha,8).getValue().toString();
  var intFogo = 0
  if(fogo.length > 0)
    var intFogo = fogo.split(", ").length;

  var terra = guiaLista.getRange(ultimalinha,9).getValue().toString();
  var intTerra = 0
  if(terra.length > 0)
    var intTerra = terra.split(", ").length;

  var metal = guiaLista.getRange(ultimalinha,10).getValue().toString();
  var intMetal = 0
  if(metal.length > 0)
    var intMetal = metal.split(", ").length;

  var agua = guiaLista.getRange(ultimalinha,11).getValue().toString();
  var intAgua = 0
  if(agua.length > 0)
    var intAgua = agua.split(", ").length;

  var listaIntElementos=[{"elem":"madeira", "val":intMadeira, "name":"Madeira"}, {"elem":"fogo","val":intFogo, "name":"Fogo"}, {"elem":"terra","val":intTerra, "name":"Terra"},{"elem":"metal","val":intMetal, "name":"Metal"}, {"elem":"agua","val":intAgua, "name":"Água"}];

  return listaIntElementos;
}

//---------------------------------------------------------------- 
//Classificar Elementos
function classifyElements(listaIntElementos){
  listaIntElementos.sort(function(a, b) {
    if(a.val > b.val) {
      return -1;
    } else {
      return true;
    }
  });

  var listaRanqueados = [];
  //var listaSegundoMaiores = [];
  var maior = listaIntElementos[0];
  
  //Primeiro
  for(i=0;i<listaIntElementos.length;i++){    
    if(maior.val == listaIntElementos[i].val){
      listaIntElementos[i]["rank"] = 1;
      listaRanqueados.push(listaIntElementos.shift());
      //listaIntElementos.pop(i);
      i-=1
    }    
  }

  //Segundo
  maior = listaIntElementos[0];
  for(i=0;i<listaIntElementos.length;i++){
    if(maior.val == listaIntElementos[i].val)
      listaIntElementos[i]["rank"] = 2;
      listaRanqueados.push(listaIntElementos.shift());
      i-=1
  }

  //Terceiro
  for(i=0;i<listaIntElementos.length;i++){    
    listaIntElementos[i]["rank"] = 3;
    listaRanqueados.push(listaIntElementos.shift());
  }  
  //var resultBigger = {"bigger":listaMaiores, "secondBiggest":listaSegundoMaiores};
  return listaRanqueados;
}

//---------------------------------------------------------------- 
//Monta os dados para envio
function buildInformationToSend(listaRanqueados){

  //para cliente e amdin
  var nomeCliente = guiaLista.getRange(ultimalinha,2).getValue(); //Range Linha x Coluna  usado em ambos  
  var adminDataCliente = guiaLista.getRange(ultimalinha,3).getValue();
  var emailCliente = guiaLista.getRange(ultimalinha,4).getValue(); //Será usado nos dois métodos de envio de email
  var adminNotaCliente = guiaLista.getRange(ultimalinha,5).getValue();
  var adminFraseCliente = guiaLista.getRange(ultimalinha,6).getValue();
  
  adminDataCliente = convertDate(adminDataCliente); 

  var countRankOne = 0;
  var countRankTwo = 0;
  var destaques = "";
  var destaque_secundario = "";  
  
  for(i=0;i<listaRanqueados.length;i++){
      if(listaRanqueados[i].rank == 1){ //Primeiros
        var class_name = listaRanqueados[i].elem;
        var body_name =  listaRanqueados[i].name;
        destaques += '<tr><td><div class="column-element-'+class_name+'"><div class="elemento-resultado">'+ body_name +'</div></div></td></tr> '; 
        countRankOne += 1;
      }else if(listaRanqueados[i].rank == 2){//Segundos   
        var class_name = listaRanqueados[i].elem;
        var body_name =  listaRanqueados[i].name;
        destaque_secundario += '<tr><td><div class="column-element-'+class_name+'"><div class="elemento-resultado">'+ body_name +'</div></div></td></tr> ';
        countRankTwo += 1;
      }      
  }

  var frase_elemento_destaque = "Seu elemento predominante"
  if(countRankOne > 1)
    frase_elemento_destaque = "Seus elementos predominantes"
  
  var frase_elemento_destaque_secundario = "Não há elemento secundário";
  if(countRankTwo > 1)
    frase_elemento_destaque_secundario = "Seus elementos secundários";
  else if (countRankTwo == 1){
    frase_elemento_destaque_secundario = "Seu elemento secundário";
  }
  
  
  var intMadeira = listaRanqueados.find( elements => elements.elem === 'madeira' ).val;
  var intFogo = listaRanqueados.find( elements => elements.elem === 'fogo' ).val;
  var intTerra = listaRanqueados.find( elements => elements.elem === 'terra' ).val;
  var intMetal = listaRanqueados.find( elements => elements.elem === 'metal' ).val;
  var intAgua = listaRanqueados.find( elements => elements.elem === 'agua' ).val;

  //Todos os Elementos para o front tanto para cliente quanto para admin

  var elementos = {
      madeira_:intMadeira,
      fogo_:intFogo,
      terra_:intTerra,
      metal_:intMetal,
      agua_:intAgua,
      frase_elemento_destaque_:frase_elemento_destaque,
      elemento_destaque_:destaques, 
      frase_elemento_destaque_secundario_:frase_elemento_destaque_secundario,
      elemento_destaque_secundario_:destaque_secundario,
      nome_client_:nomeCliente, 
      data_client_:adminDataCliente,
      email_client_:emailCliente,
      nota_client_:adminNotaCliente,
      frase_client_:adminFraseCliente
  }

  //Complementares 
  buildComplementaryElements(elementos);

  return elementos;
}
//---------------------------------------------------------------- 
//Enviar email
function sendEmail(elementos){
  var templ = HtmlService.createTemplateFromFile('client_view'); //Pagina

//vincula js > html
templ.elementos = elementos;
var htmlbody = templ.evaluate().getContent();

//Email
var email = elementos.email_client_;//guiaLista.getRange(ultimalinha,4).getValue(); //Range Linha x Coluna ;
var subjectTitle = "Resultado do questionário Descubra seu Elemento"; 
if(statusExecution == "test"){
  email = Session.getActiveUser().getEmail();// Conta em sessão
  subjectTitle = "TESTE - Resultado do questionário Descubra seu Elemento"; 
}

var mensagem = {
  to: email,
  subject: subjectTitle,
  htmlBody: htmlbody,
  name: "Terapia dos Espaços"  
}

  Logger.log("Remaining email quota: " + MailApp.getRemainingDailyQuota);

  MailApp.sendEmail(mensagem);

  Logger.log("Email enviado");
  
  guiaLista.getRange(ultimalinha,12).setValue(new Date());

}

//Email admin - segundo email
function sendEmailAdmin(elementos){
  var templ = HtmlService.createTemplateFromFile('admin_view'); //Segunda Pagina

//vincula js > html
templ.elementos = elementos;
var htmlbody = templ.evaluate().getContent();

//Email
var email = Session.getActiveUser().getEmail();// Conta em sessão
var subjectTitle = "Resultado do questionário Descubra seu Elemento "; 
if(statusExecution == "test"){  
  subjectTitle = "TESTE - Resultado do questionário Descubra seu Elemento "; 
}

var mensagem = {
  to: email,
  subject: subjectTitle + elementos.nome_client_,
  htmlBody: htmlbody,
  name: "Terapia dos Espaços"  
}

  Logger.log("Remaining email quota: " + MailApp.getRemainingDailyQuota);

  MailApp.sendEmail(mensagem);

  Logger.log("Segundo email enviado");
  
  guiaLista.getRange(ultimalinha,13).setValue(new Date());

}

//---------------------------------------------------------------- 
//Váriaveis segunda guia - configurações complementares
function buildComplementaryElements(elementos){

    planilha = SpreadsheetApp.getActiveSpreadsheet();
    guiaDois = planilha.getSheetByName("Configuracao_Email"); //Segunda guia

    var idImagemBanner = guiaDois.getRange(2,2).getValue().toString();
    var textoInicial = guiaDois.getRange(3,2).getValue().toString();
    
    var idImagemAss = guiaDois.getRange(10,2).getValue().toString();
    var empresaTitulo = guiaDois.getRange(11,2).getValue().toString();
    var subTitulo = guiaDois.getRange(12,2).getValue().toString();
    var nomes = guiaDois.getRange(13,2).getValue().toString();
    var email_assinatura = guiaDois.getRange(14,2).getValue().toString();
    var celulares = guiaDois.getRange(15,2).getValue().toString();
    var instagram = guiaDois.getRange(16,2).getValue().toString();
    var instagramTitulo = guiaDois.getRange(17,2).getValue().toString();

    //Tags 
    var imagemBanner = '<img src="https://drive.google.com/uc?export=download&id='+ idImagemBanner +'"/>';
    var imagemAss = '<img src="https://drive.google.com/uc?export=download&id='+ idImagemAss +'"/>';
        
    elementos["imagem_banner_"] = imagemBanner;
    elementos["texto_inicial_"] = textoInicial;
    
    elementos["imagem_assinatura_"] = imagemAss;
    elementos["titulo_"] = empresaTitulo;
    elementos["sub_titulo_"] = subTitulo;
    elementos["nomes_"] = nomes;
    elementos["email_assinatura_"] = email_assinatura;
    elementos["celulares_"] = celulares;
    elementos["instagram_"] = instagram;
    elementos["instagram_titulo_"] = instagramTitulo;
}
//---------------------------------------------------------------- 
function convertDate(data){
  var spreadsheet = SpreadsheetApp.getActive(); 
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(data), timezone, "dd/MM/yyyy");
}
//---------------------------------------------------------------- 
function resquestEmailTest(){
  var ui = SpreadsheetApp.getUi();
  var email = Session.getActiveUser().getEmail();
  var response = ui.alert("Envio de email de teste","O email será enviado com os dados da última pesquisa para "+ email +".  OBS: O cliente não irá receber esse email de teste. Deseja continuar e enviar o email de teste?",ui.ButtonSet.YES_NO);
  if(response == ui.Button.YES){    
    statusExecution = "test";
    main();
    ui.alert("Emails de testes encaminhados para "+ Session.getActiveUser().getEmail());
    return;
  }
  ui.alert("Cancelado envio de email de teste");

}
//----------------------------------------------------------------
function showPreviewEmail() {
  var templ = HtmlService.createTemplateFromFile('client_view'); //Pagina

  var elementos = buildInformationToSend(classifyElements(buildElements()));
  //vincula js > html
  templ.elementos = elementos;
  var htmlbody = templ.evaluate().getContent();
  
  var ui = HtmlService.createHtmlOutput(htmlbody);

  SpreadsheetApp.getUi().showModelessDialog(ui,"Pré-visualização de email");
}

function htmlHelp() {
  var html = '<html><body><div style=" color: #4f4d52; font-family: Arial"><a href="https://html-online.com/editor" target="blank"> Clique aqui para ir para página de testar os comandos</a><br><br><a href="https://developer.mozilla.org/pt-BR/" target="blank"> Clique aqui para ir para página de ajuda para pesquisar os comandos</a></div></body></html>';
  var ui = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModelessDialog(ui,"Links de ajuda");
}

//---------------------------------------------------------------- 
function main(){  
  var elementos = buildInformationToSend(classifyElements(buildElements()));
  sendEmail(elementos);
  sendEmailAdmin(elementos);

}