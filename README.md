# Sistema Web - Gerador de Memorial Descritivo

Sistema web para geração automática de documentos de memorial descritivo, convertido do script Python original do Google Colab.

## 🚀 Instalação

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

2. Execute a aplicação:
```bash
python app.py
```

3. Acesse no navegador:
```
http://localhost:5000
```

## 📋 Funcionalidades

- **Geração de Documentos Word (.docx)**:
  - Memorial Condomínio
  - Memorial Loteamento
  - Memorial Unificação
  - Memorial Desmembramento
  - Memorial Unificação e Desmembramento
  - Memorial Resumo
  - Solicitação de Análise

- **Upload de Arquivos**: Suporte para arquivos HTML/TXT de parcelas e CivilReport

- **Geração de Planilhas Excel**: Para fração ideal (condomínios) e vértices (unificação/desmembramento)

## 🗂️ Estrutura do Projeto

```
.
├── app.py                 # Aplicação Flask principal
├── memorial_processor.py  # Módulo de processamento
├── requirements.txt       # Dependências Python
├── templates/
│   └── index.html        # Interface web
├── static/
│   ├── css/
│   │   └── style.css     # Estilos
│   ├── js/
│   │   └── main.js       # JavaScript
│   └── images/            # Logos e imagens (adicionar manualmente)
└── README.md
```

## 📝 Notas

- As imagens (logos) devem ser adicionadas na pasta `static/images/`:
  - `marca_dagua.png`
  - `logo_cabecalho.png`
  - `logo_rodape.png`

- O sistema usa sessões Flask para armazenar arquivos temporariamente

- Os arquivos gerados são salvos em diretórios temporários e disponibilizados para download

## 🔧 Configuração

Para produção, altere a `SECRET_KEY` no arquivo `app.py`:

```python
app.secret_key = 'sua-chave-secreta-aqui'
```

## 📦 Dependências Principais

- Flask 3.0.0
- python-docx 1.1.0
- beautifulsoup4 4.12.2
- pandas 2.1.3
- openpyxl 3.1.2
- pyproj 3.6.1



