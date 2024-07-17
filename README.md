![Outlook Logo](https://outlookiniciarsesion01.weebly.com/uploads/9/8/5/4/98549006/outlook_orig.jpg)

# Processo de Envio de E-mails em Massa / Gmail

Este script em Python realiza o envio de e-mails em massa a partir de informações contidas em um arquivo Excel, permitindo a inclusão opcional de anexos PDF e imagens. Ele utiliza bibliotecas como `pandas`, `smtplib`, `email` e `PySimpleGUI` para manipulação de dados, envio de e-mails e criação de interface gráfica. Abaixo, detalhamos o funcionamento de cada etapa do processo:

## Bibliotecas Utilizadas

- `pandas`: Manipulação de dados em DataFrame.
- `smtplib`: Envio de e-mails através do protocolo SMTP.
- `email.mime`: Criação e manipulação de mensagens de e-mail.
- `PySimpleGUI`: Criação de interface gráfica para interação com o usuário.

## Interface Gráfica

A interface gráfica é criada utilizando o `PySimpleGUI`, com os seguintes componentes:

1. **Entrada de Caminho do Arquivo Excel**: Campo obrigatório para selecionar o arquivo Excel contendo os dados dos destinatários.
2. **Entrada de Caminho para PDF e Imagem**: Campos opcionais para anexar um PDF ou imagem ao e-mail.
3. **Assunto e Texto do E-mail**: Campos para definir o assunto e o corpo do e-mail.
4. **Botões de Ação**: Botões para enviar os e-mails, pausar ou parar o processo.

## Fluxo de Execução

1. **Leitura do Arquivo Excel**: O script lê o arquivo Excel selecionado e armazena os dados em um DataFrame.
2. **Login no Servidor SMTP**: Utiliza o servidor SMTP do Gmail para autenticação e envio dos e-mails.
3. **Iteração sobre os Destinatários**: Para cada linha no DataFrame, verifica-se o status do e-mail:
    - Se já foi enviado (`STATUS` = 'ENVIADO'), pula para o próximo.
    - Caso contrário, cria a mensagem de e-mail com os dados do destinatário e anexa arquivos, se especificados.
4. **Envio do E-mail**: Envia o e-mail para o destinatário e atualiza o DataFrame com o status e a data/hora do envio.
5. **Salvamento dos Dados Atualizados**: Salva o DataFrame atualizado com os novos status e datas de envio em um novo arquivo Excel.

## Manuseio de Anexos

- **PDF**: Se um arquivo PDF for especificado, ele é lido e anexado ao e-mail.
- **Imagem**: Similar ao PDF, a imagem é lida e anexada se especificada.

## Finalização

Após o envio de todos os e-mails, o script exibe uma mensagem de conclusão e fecha a janela da interface gráfica.

## Requisitos

- Arquivo Excel contendo colunas `nome`, `sobrenome`, `email`.
- Configuração prévia das credenciais do Gmail no script para login no servidor SMTP.

## Observações

- O script utiliza `smtplib.SMTP_SSL` para conexões seguras ao servidor Gmail.
- A interface é simplificada para facilitar o uso, mas pode ser expandida conforme necessário.
- É essencial garantir que as colunas `nome`, `sobrenome` e `email` estejam presentes no arquivo Excel para o correto funcionamento do script.
