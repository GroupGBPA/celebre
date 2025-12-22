## Processo do RPA da Célebre

Para acelerar o dia a dia da Célebre estaremos implementando um RPA que verifica se o corretor enviou um email de cotação para o cliente, coleta a PDF e anexa no perfil do cliente dentro do salesforce

### O que o RPA precisa para funcionar

1. Python 3.11 para cima
2. As bibliotecas contidas dentro do arquivo requirements.txt
3. A conta do RPA logada no Outlook e o aplicativo precisa estar _SEMPRE_ aberto
4. Chromedriver sempre atualizado
5. Login do salesforce ativo

### Funcionamento do RPA

1. Começamos iniciando a conexão com o outlook
2. Verificamos se na inbox temos algum e-mail não lido <br>
   a. Caso tenha segue o processo normal <br>
   b. Caso não tenha o processo é encerrado
3. Verificamos se o e-mail possui um anexo em .PDF, caso não tenha ignoramos o email
4. Com o PDF da cotação e o Email do cliente abrimos o Salesforce
5. Com o Salesforce logado buscar o email do cliente
6. Após abrir a página do cliente dentro do salesforce ir até o botão que anexa as cotações e subir o PDF
7. Salvar o email e o PDF no banco de dados com a data de processamento
