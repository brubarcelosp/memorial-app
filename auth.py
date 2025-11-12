"""
Módulo de autenticação com Google OAuth
Permite acesso apenas para emails @solido.arq.br e paulo.vicente001@gmail.com
"""
from flask import redirect, url_for, session, request
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from google.auth.transport import requests
from google.oauth2 import id_token
import os

# Emails permitidos
EMAILS_PERMITIDOS = [
    'paulo.vicente001@gmail.com'
]

DOMINIO_PERMITIDO = '@solido.arq.br'

# Cliente ID do Google OAuth (será configurado via variável de ambiente)
GOOGLE_CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '')

class Usuario(UserMixin):
    """Classe de usuário para Flask-Login"""
    def __init__(self, email, nome, foto_url=None):
        self.id = email
        self.email = email
        self.nome = nome
        self.foto_url = foto_url

def verificar_email_permitido(email):
    """
    Verifica se o email está na lista de emails permitidos
    ou pertence ao domínio permitido
    """
    if not email:
        return False
    
    email = email.lower().strip()
    
    # Verifica se está na lista de emails específicos
    if email in EMAILS_PERMITIDOS:
        return True
    
    # Verifica se pertence ao domínio permitido
    if email.endswith(DOMINIO_PERMITIDO.lower()):
        return True
    
    return False

def verificar_token_google(token):
    """
    Verifica o token ID do Google e retorna as informações do usuário
    """
    try:
        if not GOOGLE_CLIENT_ID:
            raise ValueError('GOOGLE_CLIENT_ID não configurado')
        
        # Verifica o token com o Google
        idinfo = id_token.verify_oauth2_token(
            token, 
            requests.Request(), 
            GOOGLE_CLIENT_ID
        )
        
        # Verifica se o token é válido e do issuer correto
        if idinfo['iss'] not in ['accounts.google.com', 'https://accounts.google.com']:
            raise ValueError('Issuer incorreto.')
        
        # Extrai informações do usuário
        email = idinfo.get('email')
        nome = idinfo.get('name', 'Usuário')
        foto_url = idinfo.get('picture')
        
        return {
            'email': email,
            'nome': nome,
            'foto_url': foto_url
        }
    except ValueError as e:
        print(f"Erro ao verificar token: {e}")
        return None

def configurar_login_manager(app):
    """Configura o Flask-Login"""
    login_manager = LoginManager()
    login_manager.init_app(app)
    login_manager.login_view = 'login'
    login_manager.login_message = 'Por favor, faça login para acessar o sistema.'
    login_manager.login_message_category = 'info'
    
    @login_manager.user_loader
    def carregar_usuario(email):
        """Carrega o usuário da sessão"""
        if 'user_email' in session and 'user_name' in session:
            return Usuario(
                session['user_email'],
                session['user_name'],
                session.get('user_picture')
            )
        return None
    
    return login_manager

