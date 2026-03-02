# Fichas de Manutencao

Aplicacao Flask + frontend estatico para consulta e operacao de fichas de manutencao.

## Estrutura do projeto

- `app.py`: backend Flask (API + entrega de arquivos estaticos)
- `frontend/`: telas HTML/CSS/JS da aplicacao
- `backend/`: scripts e dependencias Python auxiliares
- `database/`: scripts SQL de schema e verificacao
- `tools/`: utilitarios de automacao local
- `reports/`: artefatos de analise e mapeamento
- `docs/`: notas e documentacao de projeto
- `archive/`: arquivos legados preservados para historico

## Como rodar localmente

1. Criar/ativar ambiente virtual Python (`.venv`)
2. Instalar dependencias:
   - `pip install -r backend/requirements.txt`
3. Iniciar servidor:
   - `start_server.bat`
4. Acessar:
   - `http://localhost:5000/frontend/login.html`

## Observacoes

- Arquivos legados da raiz foram movidos para `archive/legacy_root_ui/`.
- A rota `/` agora abre `frontend/login.html`.
