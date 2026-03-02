# Tooling_Production

Sistema para gestao de ferramental de producao, com backend Flask, frontend estatico e integracao com banco MariaDB/MySQL.

## Objetivo

Centralizar consulta e operacao de fichas de manutencao em uma interface web local, com API para leitura de dados e exportacao.

## Stack tecnica

- Python 3.x
- Flask + Flask-CORS
- mysql-connector-python
- openpyxl
- Frontend HTML/CSS/JS
- GitHub Enterprise (GE)

## Estrutura do repositorio

- `app.py`: aplicacao Flask principal (rotas API + entrega de arquivos estaticos)
- `frontend/`: telas da aplicacao
- `frontend/assets/`: arquivos de apoio (imagens, xlsx, pdf)
- `backend/requirements.txt`: dependencias Python
- `backend/read_excel.py`: script auxiliar para leitura da planilha
- `database/`: scripts SQL
- `tools/`: scripts utilitarios locais
- `reports/`: relatorios e mapeamentos tecnicos
- `docs/notes/`: anotacoes de operacao
- `archive/legacy_root_ui/`: versao legada da UI da raiz (historico)
- `start_server.bat`: inicializacao local do servidor
- `.env.example`: exemplo de variaveis de ambiente
- `.gitignore`: arquivos/pastas ignorados no versionamento

## O que ja foi feito nesta migracao

- Commit inicial do projeto no repositorio remoto da GE.
- Integracao com historico remoto existente (`README` antigo).
- Resolucao de conflito de merge no `README.md`.
- Estrutura organizada por dominios (`frontend`, `backend`, `database`, `tools`, `docs`, `archive`).
- Rota `/` ajustada para abrir `frontend/login.html`.

## Setup local

1. Criar/ativar virtualenv:
```powershell
.\.venv\Scripts\Activate.ps1
```

2. Instalar dependencias:
```powershell
pip install -r backend\requirements.txt
```

3. Configurar `.env` baseado em `.env.example`.

## Execucao

Iniciar servidor:
```powershell
.\start_server.bat
```

Acessar:
- `http://localhost:5000/frontend/login.html`
- `http://localhost:5000/health`

## Fluxo Git diario (importante)

`git pull` baixa e integra a versao atual do remoto.  
`git push` envia seus commits locais para o remoto.

### Inicio do dia

1. Atualizar branch local:
```powershell
git checkout main
git pull origin main
```

2. Criar branch de trabalho:
```powershell
git checkout -b feature/nome-da-feature
```

### Durante o desenvolvimento

```powershell
git add .
git commit -m "feat: descricao da alteracao"
```

### Final do dia (publicar)

```powershell
git push -u origin feature/nome-da-feature
```

Depois abrir Pull Request no GitHub da GE de `feature/...` para `main`.

## Fluxo rapido (se voce trabalhar direto na main)

```powershell
git pull origin main
git add .
git commit -m "chore: atualizacao"
git push origin main
```

## Convencao de commits

- `feat:` nova funcionalidade
- `fix:` correcao de bug
- `refactor:` refatoracao sem mudar regra de negocio
- `chore:` tarefa tecnica/manutencao
- `docs:` documentacao

## Troubleshooting

### Erro de autenticacao Git

```powershell
git config --global --unset credential.helper
git config --global credential.helper manager
```

### Conflito de merge

1. Rodar `git status`
2. Editar arquivos com conflito
3. Finalizar:
```powershell
git add .
git commit -m "merge: resolve conflicts"
```

### Servidor nao sobe

1. Verificar se `.venv` existe
2. Reinstalar dependencias
3. Testar:
```powershell
python app.py
```

## Boas praticas

- Nao versionar `.env` com segredo real.
- Evitar commit de dumps pesados e logs locais.
- Fazer commits pequenos e frequentes.
- Sempre atualizar (`pull`) antes de publicar (`push`).
