import pandas as pd
from O365 import Account, FileSystemTokenBackend
import argparse

def filter_valid_emails(df):
    df = df[df['Email'].notnull()]
    df = df[df['Email'].str.contains('@', na=False)]
    return df

def main():
    parser = argparse.ArgumentParser(description="Envía emails personalizados en lote usando O365")
    parser.add_argument('--n', type=int, default=10, help='Número de emails a enviar')
    parser.add_argument('--subject', type=str, required=True, help='Asunto del email')
    parser.add_argument('--template', type=str, default='email_template.html', help='Archivo HTML de plantilla')
    parser.add_argument('--start', type=int, default=0, help='Índice inicial (base 0) de la fila en contacts.tsv')
    parser.add_argument('--client-id', type=str, required=True, help='Azure App client_id')
    parser.add_argument('--client-secret', type=str, required=True, help='Azure App client_secret')
    args = parser.parse_args()

    # Leer contactos TSV
    contacts = pd.read_csv('contacts.tsv', sep='\t')
    contacts = filter_valid_emails(contacts)

    # Seleccionar subconjunto
    contacts = contacts.iloc[args.start:args.start+args.n]

    # Leer plantilla HTML
    with open(args.template, 'r', encoding='utf-8') as f:
        html_template = f.read()

    credentials = (args.client_id, args.client_secret)
    token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
    account = Account(credentials, token_backend=token_backend)

    if not account.is_authenticated:
        # Primer uso: abre navegador para login/consentimiento
        account.authenticate(scopes=['basic', 'message_all', 'offline_access'], redirect_uri='http://localhost:8000')

    mailbox = account.mailbox()

    for idx, row in contacts.iterrows():
        html_body = html_template
        for col in contacts.columns:
            html_body = html_body.replace(f"{{{{{col}}}}}", str(row[col]))
        m = mailbox.new_message()
        m.to.add(row['Email'])
        m.subject = args.subject
        m.body = html_body
        m.body_type = 'HTML'
        m.attachments.add('signature.png')
        m.send()
        print(f"Enviado a {row['Email']} ({row['Name']})")

    print("Todos los emails han sido enviados.")

if __name__ == '__main__':
    main()
