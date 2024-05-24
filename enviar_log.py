#%%
from os import getenv

import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

import pandas as pd
#%%

load_dotenv('.env')

#%%
def email_com_anexo( destinatario, arquivo, script):

    # Arquivo com os scripts e suas utilizações
    df = pd.read_excel("arquivos/script_tabela_projeto.xlsx")
    
    #Filtrando
    df = df[df["script"] == script]
    

    # Converter o DataFrame para HTML com classes CSS
    df_html = df.to_html(index=False, classes='dataframe')

    # CSS para estilizar a tabela
    table_style = """
    <style type="text/css">
        .dataframe {
            border-collapse: collapse;
            width: 80%;
            margin: 25px 0;
            font-size: 12;
            text-align: left;
        }
        .dataframe thead th {
            background-color: #f2f2f2;
            color: #333;
            font-weight: bold;
            text-align: left;
        }
        .dataframe tbody td, .dataframe thead th {
            border: 1px solid #ddd;
            padding: 8px;
        }
        .dataframe tbody tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .dataframe tbody tr:hover {
            background-color: #f1f1f1;
        }
    </style>
    """

    # Crie a mensagem de e-mail
    smtp_server = getenv("SMTP_SERVER")
    port = getenv("SMTP_PORT")
    sender_email = getenv("SMTP_EMAIL_SENDER")
    password = getenv("SMTP_EMAIL_PASSWORD")
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = destinatario
    msg['Subject'] = "Jobs e Rotinas FP&A"


    # Corpo do e-mail em HTML
    html_content = f"""
        <html>
        <head>
        {table_style}
        </head>
        <body>
        <p>Olá,</p>
        <p>Você esta recebendo este e-mail como informativo da execução de jobs e rotinas do FP&A.</p>

        <p>A tabela abaixo descreve todos os recursos e trabalhos que <i>foram atualizados ou não</i> a partir da execução do script <b> {script}.</b>.</p>
        <p>Portanto, analise o log em anexo e identifique se existem ou não problemas de atualização.</p>
        {df_html}
        <p>Atenciosamente,<br>Equipe FP&A</p>
        </body>
        </html>
        """

    # Anexe o corpo do e-mail
    msg.attach(MIMEText(html_content, 'html'))


    # Anexe o arquivo
    with open(arquivo, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={arquivo}',
        )
        msg.attach(part)

    # Envie o e-mail
    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=context)
        server.login(sender_email, password)
        server.sendmail(sender_email, destinatario, msg.as_string())
