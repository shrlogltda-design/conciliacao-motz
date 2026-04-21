# Dashboard Conciliação MOTZ

Aplicação Streamlit que executa a skill `conciliacao-motz` diretamente no navegador.
Sua equipe sobe os arquivos (PDFs Repom, MOTZ XLSX, ATUA XLS) e recebe:

- Planilha XLSX consolidada com formatação condicional
- Dashboard interativo com KPIs, filtros de data/status e busca
- Export em CSV da tabela filtrada

## Estrutura

```
streamlit-app/
├── app.py                  # Aplicação Streamlit principal
├── requirements.txt        # Dependências Python
├── packages.txt            # Dependências de sistema (LibreOffice)
├── .streamlit/
│   └── config.toml         # Tema e limite de upload
└── scripts/
    ├── conciliacao.py      # [VOCÊ COPIA] Script principal da skill
    └── parse_xls.py        # [VOCÊ COPIA] Parser BIFF8 fallback
```

## Passo a passo para publicar no Streamlit Cloud (grátis)

### 1. Testar localmente (opcional)

```bash
cd streamlit-app
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`.

### 2. Subir para o GitHub

```bash
cd streamlit-app
git init
git add .
git commit -m "Dashboard conciliação MOTZ"
gh repo create conciliacao-motz --private --source=. --push
```

(Ou crie o repo manualmente no GitHub e faça push.)

### 3. Deploy no Streamlit Cloud

1. Acesse https://share.streamlit.io
2. Login com sua conta GitHub
3. Clique em **New app**
4. Aponte para o repo `conciliacao-motz`, branch `main`, arquivo `app.py`
5. Clique **Deploy**

Em ~3 minutos você tem uma URL tipo `https://conciliacao-motz.streamlit.app` para compartilhar com a equipe.

### 4. Controle de acesso (recomendado)

No painel do app no Streamlit Cloud, vá em **Settings → Sharing** e adicione os emails dos 2-5 usuários do seu time. Só quem estiver na lista conseguirá acessar.

## Uso

1. Abra a URL do app no navegador
2. Suba os PDFs Repom (pode selecionar vários)
3. Suba o arquivo MOTZ (.xlsx)
4. Suba o arquivo ATUA (.xls)
5. Clique em **Rodar conciliação**
6. Aguarde ~30s-2min (depende do tamanho)
7. Explore o dashboard · baixe o XLSX consolidado ou CSV filtrado

Alternativa: se você já tem a planilha consolidada gerada em outro lugar, clique em **Carregar XLSX pronto** para pular direto pro dashboard.

## Limites

- **Upload:** 200 MB por arquivo (ajustável em `.streamlit/config.toml`)
- **Execução:** 300 segundos de timeout para a conciliação
- **Streamlit Cloud grátis:** 1 GB de RAM, suficiente para volumes normais

## Troubleshooting

**"scripts/conciliacao.py não encontrado"** — Os scripts já vêm inclusos em `scripts/`. Se o erro aparecer, verifique se a pasta foi mantida no commit do Git (`git status` e reveja `.gitignore`).

**"LibreOffice não encontrado"** — O `packages.txt` já instala. Se estiver rodando local, instale via `sudo apt install libreoffice`.

**Timeout na conciliação** — Arquivos muito grandes. Aumente `timeout=300` na chamada `subprocess.run` no `app.py`.
