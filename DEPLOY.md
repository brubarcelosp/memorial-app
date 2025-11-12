# Deploy no Heroku

## Pré-requisitos

1. Conta no Heroku (https://www.heroku.com)
2. Heroku CLI instalado (https://devcenter.heroku.com/articles/heroku-cli)
3. Git instalado

## Passos para Deploy

### 1. Login no Heroku

```bash
heroku login
```

### 2. Criar aplicação no Heroku

```bash
heroku create nome-da-sua-app
```

Ou use o dashboard do Heroku para criar uma nova app.

### 3. Configurar variáveis de ambiente (se necessário)

```bash
heroku config:set SECRET_KEY=sua-chave-secreta-aqui
```

### 4. Fazer commit dos arquivos

```bash
git init
git add .
git commit -m "Initial commit"
```

### 5. Conectar ao Heroku e fazer deploy

```bash
heroku git:remote -a nome-da-sua-app
git push heroku main
```

Ou se estiver usando a branch `master`:

```bash
git push heroku master
```

### 6. Verificar logs

```bash
heroku logs --tail
```

### 7. Abrir a aplicação

```bash
heroku open
```

## Arquivos Importantes

- **Procfile**: Define como o Heroku deve iniciar a aplicação
- **runtime.txt**: Especifica a versão do Python
- **requirements.txt**: Lista todas as dependências
- **.gitignore**: Arquivos que não devem ser commitados

## Notas Importantes

1. **Porta**: O Heroku define a porta automaticamente via variável de ambiente `PORT`
2. **Debug**: Desative o modo debug em produção (já está configurado para usar `PORT`)
3. **Arquivos estáticos**: As imagens em `static/images/` devem ser commitadas
4. **Uploads**: Arquivos temporários são limpos automaticamente

## Troubleshooting

### Erro de porta
Se houver erro relacionado à porta, verifique se o código está usando `os.environ.get('PORT', 5001)`

### Erro de dependências
Verifique se todas as dependências estão no `requirements.txt`

### Erro de build
Verifique os logs com `heroku logs --tail`

### Limite de memória
O Heroku tem limites de memória. Se necessário, considere usar um dyno maior.

## Comandos Úteis

```bash
# Ver status da aplicação
heroku ps

# Ver logs em tempo real
heroku logs --tail

# Abrir console Python
heroku run python

# Reiniciar a aplicação
heroku restart

# Ver configurações
heroku config

# Escalar dynos
heroku ps:scale web=1
```

