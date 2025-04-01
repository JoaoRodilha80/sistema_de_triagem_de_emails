import imaplib
import email
import smtplib
import joblib
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from email.message import EmailMessage
from email.header import decode_header
from preprocessamento import TextPreprocessor
import config

EMAIL_USER = config.EMAIL_CONFIG["USER"]
EMAIL_PASS = config.EMAIL_CONFIG["PASS"]
IMAP_SERVER = config.EMAIL_CONFIG["IMAP_SERVER"]
SMTP_SERVER = config.EMAIL_CONFIG["SMTP_SERVER"]
CHECK_INTERVAL = config.APP_CONFIG["CHECK_INTERVAL"]

try:
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=10)
    server.login('triagem.ctrlaltesc@gmail.com', 'tpgq zcbw icmg mfax')
    print("✅ Login bem-sucedido!")
    server.quit()
except Exception as e:
    print(f"❌ Falha: {e}")


def decode_email_header(header):
    """Decodifica cabeçalhos de e-mail com caracteres especiais."""
    decoded_parts = []
    for part, encoding in decode_header(header):
        if isinstance(part, bytes):
            decoded_parts.append(part.decode(encoding or 'utf-8'))
        else:
            decoded_parts.append(str(part))
    return ' '.join(decoded_parts)

# Carregar modelo
try:
    modelo = joblib.load(config.APP_CONFIG["MODEL_FILE"])
    preprocessor = joblib.load(config.APP_CONFIG["PREPROCESSOR_FILE"])
except Exception as e:
    print(f"❌ Erro ao carregar modelo: {e}")
    exit()

def is_problem_paragraph(text):
    """Verifica se o texto descreve um problema técnico"""
    text_lower = text.lower()
    problem_keywords = [
        'problema', 'erro', 'dificuldade', 'não funciona', 
        'não está', 'falha', 'travando', 'lento', 'quebrado',
        'defeito', 'assistência', 'suporte', 'bug', 'crash',
        'não abre', 'não roda', 'não liga', 'não reconhece',
        'não responde', 'não carrega', 'não conecta',
        'piscando', 'fechando', 'fecha sozinho'
    ]
    return any(keyword in text_lower for keyword in problem_keywords)

def is_general_text(text):
    """Identifica textos genéricos que não são problemas técnicos"""
    text_lower = text.lower()
    general_phrases = [
        'olá', 'oi ', 'ola ', 'bom dia', 'boa tarde', 'boa noite',
        'atenciosamente', 'obrigado', 'grato', 'agradeço',
        'por favor', 'preciso de ajuda', 'urgente', 'quanto antes',
        'favor', 'contato', 'retorno', 'resposta', 'ajuda',
        'att', 'cordialmente', 'desde já agradeço'
    ]
    return any(phrase in text_lower for phrase in general_phrases)

def detect_multiple_issues(body):
    """Identifica parágrafos que descrevem problemas reais"""
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    paragraphs = [p.strip() for p in body.split('\n\n') if p.strip()]
    
    problem_paragraphs = []
    for para in paragraphs:
        if len(para.split()) < 5 or is_general_text(para):
            continue
            
        if is_problem_paragraph(para):
            problem_paragraphs.append(para)
    
    if len(problem_paragraphs) < 1:
        return None
    
    issues = []
    for para in problem_paragraphs:
        category = classify_email("", para)
        issues.append((para, category))
    
    return issues

def classify_email(subject, body):
    texto = f"{subject} {body}".lower()
    
    # Regras manuais prioritárias - Balanceadas
    hardware_terms = [
        'monitor', 'tela', 'piscando', 'quebrad', 'danificad', 
        'fisic', 'não liga', 'hardware', 'hd', 'ssd', 'memória', 
        'ram', 'processador', 'mouse', 'teclado', 'cabo', 'conector'
    ]
    
    software_terms = [
        'excel', 'fechando', 'programa', 'aplicativo', 'não abre', 
        'não roda', 'crash', 'travando', 'lento', 'erro', 'bug',
        'software', 'sistema', 'windows', 'instalar', 'atualização',
        'formula', 'planilha', 'word', 'powerpoint', 'aplicação'
    ]
    
    # Contagem de termos para balanceamento
    hardware_count = sum(1 for term in hardware_terms if term in texto)
    software_count = sum(1 for term in software_terms if term in texto)
    
    if hardware_count > 0 or software_count > 0:
        if hardware_count > software_count:
            return "hardware"
        elif software_count > hardware_count:
            return "software"
    
    # Se empate ou nenhum termo específico, verifica regras manuais
    if any(term in texto for term in ['lento', 'lentidão', 'travando', 'congelando']):
        return "software"
    elif any(palavra in texto for palavra in ['wifi', 'internet', 'rede']):
        if any(termo in texto for termo in ['antena', 'placa', 'fisic']):
            return "hardware"
        return "software"
    
    # Classificação pelo modelo ML
    try:
        texto_processado = preprocessor.preprocess(texto)
        return modelo.predict([texto_processado])[0]
    except Exception as e:
        print(f"⚠️ Erro no modelo: {e}")
        return "indefinido"

def connect_imap():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, timeout=30)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")
    return mail

def get_unread_emails(mail):
    _, messages = mail.search(None, "UNSEEN")
    return messages[0].split()

def extract_email_content(mail, email_id):
    try:
        _, data = mail.fetch(email_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)
        
        subject = decode_email_header(msg["Subject"]) if msg["Subject"] else "Sem assunto"
        sender = decode_email_header(msg["From"])
        
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors='ignore')
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors='ignore')
        
        return {
            "id": email_id,
            "subject": subject,
            "sender": sender,
            "body": body
        }
    except Exception as e:
        print(f"⚠️ Erro ao extrair e-mail: {e}")
        return None

def forward_email(original_email, category):
    try:
        # Adicione este timeout:
        with smtplib.SMTP_SSL(SMTP_SERVER, 465, timeout=10) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
        recipient = recipient = config.EMAIL_CONFIG["SOFTWARE_TEAM"] if category == "software" else config.EMAIL_CONFIG["HARDWARE_TEAM"]
        msg = EmailMessage()
        msg["From"] = EMAIL_USER
        msg["To"] = recipient
        msg["Subject"] = f"[{category.upper()}] {original_email['subject'][:50]}..."
        
        email_body = f"De: {original_email['sender']}\n\n"
        email_body += f"Problema relatado ({category.upper()}):\n\n"
        email_body += f"{original_email['body']}\n\n"
        email_body += f"---\nAssunto original: {original_email['subject']}"
        
        msg.set_content(email_body)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
    except Exception as e:
        print(f"⚠️ Erro ao encaminhar: {e}")

def log_to_spreadsheet(email_data, category):
    """Registra o ticket na planilha Google Sheets correspondente."""
    try:
        # Configuração da autenticação
        scope = scope = config.SHEETS_CONFIG["SCOPE"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(config.SHEETS_CONFIG["CREDENTIALS_FILE"], scope)
        client = gspread.authorize(creds)

        # Seleciona a planilha correta
        sheet_name = "Tickets_Software" if category == "software" else "Tickets_Hardware"
        sheet = client.open(sheet_name).sheet1

        # Extrai nome e e-mail do remetente
        sender_name = email_data["sender"].split("<")[0].strip()
        sender_email = email_data["sender"].split("<")[-1].replace(">", "").strip()

        # Dados a serem inseridos
        row_data = [
            sender_name,
            sender_email,
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            email_data["subject"]
        ]

        # Adiciona nova linha
        sheet.append_row(row_data)
        print(f"✅ Log adicionado na planilha {sheet_name}")

    except Exception as e:
        print(f"⚠️ Falha ao registrar na planilha: {e}")

def send_automatic_reply(original_email, category):
    try:
        sender_name = original_email["sender"].split()[0]
        sender_name = ''.join(c for c in sender_name if c.isalpha())
        
        msg = EmailMessage()
        msg["From"] = EMAIL_USER
        msg["To"] = original_email["sender"]
        msg["Subject"] = f"✔ Ticket recebido - Suporte {category.upper()}"
        
        if category == "software":
            content = f"""Olá {sender_name},

Seu ticket foi classificado como SOFTWARE e já foi encaminhado para nossa equipe especializada.

📅 Prazo de resposta: 2 horas úteis
🔍 Descrição do problema: "{original_email['subject']}"

Agradecemos pela paciência!
Equipe CtrlAltEsc"""
        elif category == "hardware":
            content = f"""Olá {sender_name},

Seu ticket foi classificado como HARDWARE e está sendo processado por nossos técnicos.

📅 Prazo de resposta: 4 horas úteis
🔧 Descrição do problema: "{original_email['subject']}"

Atenciosamente,
Equipe CtrlAltEsc"""
        else:
            content = f"""Olá {sender_name},

Recebemos seu ticket e estamos analisando a melhor equipe para atendê-lo.

📅 Prazo de resposta: 1 dia útil
📝 Descrição do problema: "{original_email['subject']}"

Obrigado por nos contatar!
Equipe CtrlAltEsc"""

        msg.set_content(content)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        
        log_to_spreadsheet(original_email, category)  # Novo log!

    except Exception as e:
        print(f"⚠️ Erro ao enviar resposta automática: {e}")

def send_multiple_issue_reply(original_email, issues):
    try:
        msg = EmailMessage()
        msg["From"] = EMAIL_USER
        msg["To"] = original_email["sender"]
        msg["Subject"] = "✔ Seus problemas técnicos foram triados"
        
        body_lines = [
            "Olá,",
            "",
            "Identificamos os seguintes problemas técnicos em seu e-mail:",
            ""
        ]
        
        for i, (text, category) in enumerate(issues, 1):
            team = "SOFTWARE" if category == "software" else "HARDWARE"
            body_lines.append(f"🔹 Problema {i} (Equipe de {team}):")
            body_lines.append(f'"{text[:150]}..."')
            body_lines.append("")
        
        body_lines.extend([
            "Prazos estimados para resolução:",
            "",
            "🖥️ SOFTWARE: Resposta em até 2 horas úteis",
            "🔧 HARDWARE: Resposta em até 4 horas úteis",
            "",
            "Agradecemos pela compreensão!",
            "Equipe CtrlAltEsc"
        ])
        
        msg.set_content("\n".join(body_lines))
        
        with smtplib.SMTP_SSL(SMTP_SERVER, 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        
        for _, category in issues:
            log_to_spreadsheet(original_email, category)  # Log para cada problema

    except Exception as e:
        print(f"⚠️ Erro ao enviar resposta: {e}")

def mark_as_read(mail, email_id):
    mail.store(email_id, "+FLAGS", "\\Seen")

def process_emails():
    try:
        mail = connect_imap()
        emails = get_unread_emails(mail)
        
        for email_id in emails:
            email_data = extract_email_content(mail, email_id)
            if not email_data:
                continue
                
            print(f"\n📩 Novo e-mail: {email_data['subject']}")
            print(f"👤 Remetente: {email_data['sender']}")
            
            print("🔍 Análise de parágrafos:")
            body = email_data['body'].replace('\r\n', '\n').replace('\r', '\n')
            paragraphs = [p.strip() for p in body.split('\n\n') if p.strip()]
            
            for i, p in enumerate(paragraphs[:5], 1):
                status = "✅ PROBLEMA" if is_problem_paragraph(p) else "⏭️ IGNORADO"
                print(f"   {i}. [{status}] {p[:80]}{'...' if len(p) > 80 else ''}")
            
            multiple_issues = detect_multiple_issues(email_data["body"])
            
            if multiple_issues and len(multiple_issues) >= 2:
                print("🔧 Múltiplos problemas detectados:")
                for i, (text, category) in enumerate(multiple_issues, 1):
                    print(f"   {i}. [{category.upper()}] {text[:80]}...")
                
                for text, category in multiple_issues:
                    partial_email = {
                        "id": email_data["id"],
                        "subject": f"[PARTE {category.upper()}] {email_data['subject']}",
                        "sender": email_data["sender"],
                        "body": text
                    }
                    forward_email(partial_email, category)
                
                send_multiple_issue_reply(email_data, multiple_issues)
            else:
                category = classify_email(email_data["subject"], email_data["body"])
                print(f"🔍 Categoria única: {category}")
                forward_email(email_data, category)
                send_automatic_reply(email_data, category)
            
            mark_as_read(mail, email_id)
        
        mail.close()
        mail.logout()
    except Exception as e:
        print(f"⚠️ Erro fatal: {e}")

if __name__ == "__main__":
    print("🔄 Iniciando monitoramento de e-mails...")
    while True:
        process_emails()
        print(f"⏳ Aguardando {CHECK_INTERVAL} segundos...")
        time.sleep(CHECK_INTERVAL)