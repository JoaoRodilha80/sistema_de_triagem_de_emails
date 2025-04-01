# config.example.py - Arquivo de configuração de exemplo
# Renomeie para config.py e preencha com seus dados reais

# Configurações de Email (substitua com suas credenciais)
EMAIL_CONFIG = {
    "USER": "seu_email@gmail.com",          # Substitua pelo e-mail real
    "PASS": "sua_senha_aqui",              # Substitua pela senha real
    "IMAP_SERVER": "imap.gmail.com",       # Servidor IMAP (Gmail)
    "SMTP_SERVER": "smtp.gmail.com",       # Servidor SMTP (Gmail)
    "SOFTWARE_TEAM": "email_software@exemplo.com",  # E-mail da equipe de software
    "HARDWARE_TEAM": "email_hardware@exemplo.com"   # E-mail da equipe de hardware
}

# Configurações do Google Sheets (substitua com seu arquivo de credenciais)
SHEETS_CONFIG = {
    "CREDENTIALS_FILE": "credenciais.json",  # Arquivo baixado do Google Cloud
    "SCOPE": [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
}

# Configurações da Aplicação
APP_CONFIG = {
    "CHECK_INTERVAL": 10,                  # Tempo entre verificações (em segundos)
    "MODEL_FILE": "modelo_classificador.pkl",  # Arquivo do modelo de ML
    "PREPROCESSOR_FILE": "preprocessor.pkl"    # Arquivo do pré-processador de texto

}