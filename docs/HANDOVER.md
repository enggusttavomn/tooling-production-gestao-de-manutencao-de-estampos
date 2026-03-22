# Handover Tecnico

Este guia resume os pontos principais para quem vai assumir a manutencao do sistema.

## 1. Visao geral

- API principal em `backend/api/app.py`.
- Frontend estatico em `frontend/`.
- SQL e dados de banco em `database/`.
- Documentacao de operacao em `docs/`.

## 2. Execucao local

1. Ativar virtualenv `.venv` na raiz.
2. Instalar dependencias com `pip install -r backend/requirements.txt`.
3. Criar `.env` a partir de `.env.example`.
4. Iniciar com `backend/scripts/start_server.bat`.
5. Validar `http://localhost:5000/health`.

## 3. Scripts operacionais

- `backend/scripts/start_server.bat`: inicia API e tenta subir MariaDB local.
- `backend/scripts/start_local_mariadb.ps1`: start da instancia local MariaDB.
- `backend/scripts/stop_local_mariadb.ps1`: stop da instancia local.
- `backend/scripts/import_dump_local.ps1`: importa dump SQL removendo definers.
- `backend/scripts/watch_excel.ps1`: monitora planilha fonte e regenera HTML.

## 4. Contrato de API (resumo)

- `GET /api/fichas`: lista com filtros e cache curto.
- `GET /api/ficha/<id>`: detalhe de ficha.
- `GET /api/estatisticas`: indicadores para dashboard.
- `GET /api/exportar-excel`: exportacao XLSX.
- `POST/PUT/DELETE`: endpoints bloqueados no modo atual (somente leitura).

## 5. Convencoes de manutencao

- Evitar regra de negocio no frontend.
- Comentarios devem explicar contexto/decisao.
- Modulos Python devem manter docstrings em modulo/funcoes publicas.
- Scripts utilitarios devem ser executaveis de forma isolada.

## 6. Higiene de repositorio

- Nao versionar logs e artefatos temporarios.
- Logs locais ficam em `backend/logs/`.
- Temporarios de scripts ficam em `backend/tmp/`.
- Atualizar `.gitignore` se surgirem novos artefatos locais.

## 7. Pontos de atencao

- `backend/api/app.py` ainda concentra responsabilidades de API e arquivos estaticos.
- Revisar performance periodicamente no endpoint `GET /api/fichas`.

## 8. Proximo passo recomendado

- Modularizar `backend/api/app.py` em componentes (`db`, `services`, `routes`) mantendo contrato atual da API.
