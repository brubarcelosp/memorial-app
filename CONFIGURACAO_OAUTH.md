# 🔐 Configuração de Autenticação Google OAuth

## Passo 1: Criar Credenciais no Google Cloud Console

1. Acesse: https://console.cloud.google.com/
2. Crie um novo projeto ou selecione um existente
3. Vá em **APIs & Services** → **Credentials**
4. Clique em **Create Credentials** → **OAuth client ID**
5. Se solicitado, configure a tela de consentimento OAuth:
   - Tipo: **External** (ou Internal se tiver Google Workspace)
   - Nome do app: "Gerador de Memorial Descritivo"
   - Email de suporte: seu email
   - Adicione seu email como test user (se for External)
6. Configure o OAuth client:
   - Tipo: **Web application**
   - Nome: "Memorial App Web Client"
   - **Authorized JavaScript origins**: 
     - `http://localhost:5001` (desenvolvimento)
     - `https://seu-dominio.com` (produção)
   - **Authorized redirect URIs**:
     - `http://localhost:5001/login` (desenvolvimento)
     - `https://seu-dominio.com/login` (produção)
7. Copie o **Client ID** gerado

## Passo 2: Configurar Variáveis de Ambiente

### Desenvolvimento Local

Crie um arquivo `.env` na raiz do projeto:

```bash
GOOGLE_CLIENT_ID=seu-client-id-aqui.apps.googleusercontent.com
SECRET_KEY=sua-chave-secreta-aleatoria-aqui
```

Ou exporte as variáveis no terminal:

```bash
export GOOGLE_CLIENT_ID="seu-client-id-aqui.apps.googleusercontent.com"
export SECRET_KEY="sua-chave-secreta-aleatoria-aqui"
```

### Produção (Heroku, etc.)

Configure as variáveis de ambiente na plataforma:

```bash
heroku config:set GOOGLE_CLIENT_ID="seu-client-id-aqui"
heroku config:set SECRET_KEY="sua-chave-secreta-aleatoria-aqui"
```

## Passo 3: Emails Permitidos

Os seguintes emails têm acesso ao sistema:
- `paulo.vicente001@gmail.com`
- Qualquer email do domínio `@solido.arq.br`

Para adicionar mais emails, edite o arquivo `auth.py`:

```python
EMAILS_PERMITIDOS = [
    'paulo.vicente001@gmail.com',
    'outro-email@gmail.com'  # Adicione aqui
]

DOMINIO_PERMITIDO = '@solido.arq.br'
```

## Passo 4: Testar

1. Inicie o servidor:
   ```bash
   python app.py
   ```

2. Acesse: http://localhost:5001/login

3. Clique em "Entrar com Google"

4. Faça login com um email permitido

5. Você será redirecionado para a página principal

## Troubleshooting

### Erro: "Google Client ID não configurado"
- Verifique se a variável `GOOGLE_CLIENT_ID` está definida
- Reinicie o servidor após definir a variável

### Erro: "Acesso negado"
- Verifique se o email está na lista de permitidos
- Verifique se o email termina com `@solido.arq.br`

### Erro: "Token inválido"
- Verifique se o Client ID está correto
- Verifique se as URLs autorizadas no Google Console estão corretas
- Certifique-se de que está usando HTTPS em produção

## Segurança

- ✅ Nunca commite o `.env` ou credenciais no Git
- ✅ Use uma `SECRET_KEY` forte e aleatória em produção
- ✅ Configure HTTPS em produção
- ✅ Mantenha a lista de emails permitidos atualizada

