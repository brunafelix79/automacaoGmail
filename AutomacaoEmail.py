import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import datetime
import PySimpleGUI as sg

# Configurações iniciais
sg.theme('Reddit')

# Variáveis para armazenar os caminhos dos arquivos
excel_path = None
pdf_path = None
img_path = None

layout = [
    [sg.Text('Excel.:*'), sg.Input(key='excel_path', size=(61, 1)), sg.FileBrowse()],
    [sg.Text('Possui PDF?.:'), sg.Input(key='possui_pdf', size=(55, 1)), sg.FileBrowse()],
    [sg.Text('Possui Imagem?.:'), sg.Input(key='possui_imagem', size=(52, 1)), sg.FileBrowse()],
    [sg.Text('Assunto.:*'), sg.Input(key='assunto', size=(58, 1))],
    [sg.Text('Texto.:*'), sg.Multiline(key='texto', size=(68, 15))],
    [sg.Button('Enviar'), sg.Button('Pause'), sg.Button('Parar', button_color=('white', 'red'))]
]

# Cria a janela principal
janela = sg.Window('ENVIO EMAIL EM MASSA - GMAIL', layout)

while True:
    event, values = janela.read()

    if event == sg.WIN_CLOSED:
        janela.close()
        break

    elif event == 'Enviar':
        if not values['excel_path']:
            sg.popup('Por favor, selecione um arquivo Excel.', title='Erro')
            continue
        excel_path = values['excel_path']
        if values['possui_pdf']:
            pdf_path = values['possui_pdf']
        if values['possui_imagem']:
            img_path = values['possui_imagem']

        # Lê o arquivo Excel
        df = pd.read_excel(excel_path)
        n_rows = len(df.index)
        progresso = 0

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            # Faz login no servidor SMTP
            server.login("", "")

            # Itera sobre as linhas do DataFrame
            for index, row in df.iterrows():
                if row.get('STATUS') == 'ENVIADO':
                    continue  # Pula os e-mails que já foram enviados

                # Cria o objeto MIMEMultipart para o e-mail
                msg = MIMEMultipart()
                msg['From'] = ""

                nome = row.get('nome')
                sobrenome = row.get('sobrenome')
                email = row.get('email')

                if 'nome' not in df.columns or 'sobrenome' not in df.columns or 'email' not in df.columns:
                    raise KeyError("As colunas 'nome', 'sobrenome' e 'email' não estão presentes no arquivo Excel.")

                msg['To'] = email
                msg['Subject'] = values['assunto']

                body = f'Prezados(as) {nome} {sobrenome},\n{values["texto"]}' if values['texto'] else ""
                msg.attach(MIMEText(body, 'plain'))

                if pdf_path:
                    filename = pdf_path
                    with open(filename, "rb") as attachment:
                        part = MIMEApplication(attachment.read())
                        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                        msg.attach(part)

                # Envia o e-mail
                text = msg.as_string()
                server.sendmail("seuemail@gmail.com", email, text)
                print(f"E-mail enviado para {nome} {sobrenome} ({email}) ")

                # Atualiza a barra de progresso
                progresso += 1

                # Atualiza o status e a data e hora de envio no DataFrame df
                df.at[index, 'STATUS'] = 'ENVIADO'
                df.at[index, 'DATA_HORA_ENVIO'] = datetime.datetime.now()

                # Salva o DataFrame df com as novas colunas de status e data e hora de envio como um novo arquivo Excel
                novo_excel_path = f'C:/TESTES/ENVIADOS_{datetime.datetime.now().strftime("%d_%m_%y_%H_%M_%S")}.xlsx'
                df.to_excel(novo_excel_path, index=False)

        # Exibe uma mensagem ao usuário quando o envio estiver completo
        sg.popup("Finalizado!", auto_close=True, auto_close_duration=500)

# Fecha a janela de envio
janela.close()

