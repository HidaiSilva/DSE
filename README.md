# DSE

Descubra Seu Elemento - projeto de automatização de criação e envio de emails dos resultados de pesquisa realizado via Google Form.

----------------
## Execução Inicial

- Abrir uma planilha vinculada ao Form; 
- Aba vinculada ao Form deve ser renomeada para “Respostas”;
- Criar uma segunda aba com chamada “Configuracao_Email”;
- Através da planilha abra o App Script;
- Na aba editor crie três arquivos [client.gs], [admin_view.html] e [client_view.html];
- Copie os códigos disponíveis aqui no git para os arquivos criados no App Script, conforme seus respectivos nomes e salve o projeto;  
- Abra o [client.gs]  na função “setUpTrigger” o “.forForm()” deve receber o id do formulário.

~~~JS
//Trigger Formulario
function setUpTrigger() {
  ScriptApp.newTrigger('main')
  .forForm('ID DO FORMULÁRIO')
  .onFormSubmit()
  .create();
} 
~~~

- Ainda no [client.gs], na opção de selecionar função para executar selecione  “setUpTrigger” e clique em executar;

## Configuração de variáveis disponíveis na planilha aba Respostas

- Na planilha na aba “Respostas”, devesse criar os campos que serão apresentados no frontend/email como os resultados da pesquisa.

  ![Aba Resposta](images\img_aba_respostas.PNG) 

- A atenção precisa ser dada nas posições dos campos, pois irão ser chamados pelo  [client.gs] nas funções “buildElements”, “buildComplementaryElements” e “buildInformationToSend” onde é realizado um range linha x coluna para saber exatamente quais campos pegar. 

## Configuração de variáveis disponíveis na planilha aba Configuracao_Email

- Na planilha na aba “Configuracao_Email”, devesse criar os campos que serão apresentados no frontend/email, tais como descrição no topo e assinatura do email.

![Aba Configuração Email](images\img_aba_configuracao_email.PNG) 

- Novamente a atenção precisa ser dada nas posições dos campos, pois irão ser chamados pelo  [client.gs] nas funções “buildElements”, “buildComplementaryElements” e “buildInformationToSend” onde é realizado um range linha x coluna para saber exatamente quais campos pegar. 
### Configuração dos botões 
	exemplo imagem
------------------
## Fluxo de processo

Imagem


[client.gs]: https://github.com/HidaiSilva/DSE/blob/main/gs/client.gs

[client_view.html]: https://github.com/HidaiSilva/DSE/blob/main/view/client_view.html

[admin_view.html]: https://github.com/HidaiSilva/DSE/blob/main/view/admin_view.html

