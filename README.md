# Tooling Production — Fichas de Manutencao de Estampos

## Sumario

1. [O que e o sistema](#1-o-que-e-o-sistema)
2. [Stack tecnologica](#2-stack-tecnologica)
3. [Estrutura de pastas](#3-estrutura-de-pastas)
4. [Perfis de usuario e fluxo de acesso](#4-perfis-de-usuario-e-fluxo-de-acesso)
5. [Mapa de telas](#5-mapa-de-telas)
6. [Setup local do zero](#6-setup-local-do-zero)
7. [Variaveis de ambiente](#7-variaveis-de-ambiente)
8. [Como iniciar o servidor](#8-como-iniciar-o-servidor)
9. [Endpoints de API](#9-endpoints-de-api)
10. [Banco de dados](#10-banco-de-dados)
11. [Scripts operacionais](#11-scripts-operacionais)
12. [Fluxo Git](#12-fluxo-git)
13. [Convencoes de commit](#13-convencoes-de-commit)
14. [Troubleshooting](#14-troubleshooting)
15. [Roadmap e proximos passos](#15-roadmap-e-proximos-passos)

---

## 1. O que e o sistema

O **Fichas de Manutencao de Estampos** e um sistema web interno da GE Vernova para registrar, acompanhar e consultar fichas de manutencao de ferramentais (estampos) utilizados na producao.

Em termos praticos: ele substitui o controle espalhado em planilhas e organiza o ciclo da ficha por perfil, deixando cada area preencher apenas a parte que lhe cabe.

### Problema que resolve

Antes do sistema, as fichas eram controladas manualmente em planilhas Excel, dificultando rastreabilidade, historico e visibilidade de pendencias entre turnos. O sistema centraliza tudo no banco de dados da unidade e disponibiliza telas especificas por funcao.

### Quem usa

| Perfil | Responsabilidade |
|---|---|
| **Producao** | Abre a ficha quando identifica necessidade de manutencao |
| **Ferramenteiro** | Preenche os dados de manutencao executada |
| **Engenheiro de Processo** | Valida, aprova e acompanha via dashboard e relatorios |
| **Planejamento** | Consulta o roteiro de JOBs e plano de manutencao de estampos |
| **Admin** | Acesso total a todas as telas e funcionalidades |

### Fluxo operacional de ponta a ponta

1. **Producao** identifica uma necessidade de manutencao e abre a ficha.
2. A ficha passa a aparecer nas listas de pendencia das outras areas.
3. **Ferramenteiro** preenche as informacoes de execucao da manutencao.
4. **Engenharia** revisa, acompanha indicadores e consulta historico.
5. **Planejamento** usa as telas de planejamento para cruzar informacoes de JOB, ferramental e condicao dos estampos.
6. **Admin** pode navegar por todo o fluxo e atuar como apoio operacional.

---

## 2. Stack tecnologica

| Camada | Tecnologia | Versao |
|---|---|---|
| Backend / API | Python + Flask | Flask 3.0.0 |
| Banco de dados | MariaDB (servidor GE) | — |
| Conector DB | mysql-connector-python | 9.1.0 |
| Exportacao Excel | openpyxl | 3.11.2 |
| Variaveis de ambiente | python-dotenv | 1.0.1 |
| Frontend | HTML5 + CSS3 + JavaScript (vanilla) | — |
| CORS | Flask-CORS | 4.0.0 |
| Servidor local | MariaDB portable (.mariadb_local/) | — |

Nao ha framework frontend (sem React, Vue etc.). Tudo e HTML/CSS/JS puro servido estaticamente pelo proprio Flask.

---

## 3. Estrutura de pastas

```
FICHAS_DE_MANUTENCAO/
│
├── .env                        # Credenciais reais (nunca versionar)
├── .env.example                # Template de variaveis de ambiente
├── .gitignore                  # Exclui .env, logs, __pycache__, .venv etc.
├── README.md                   # Este arquivo
│
├── backend/
│   ├── api/
│   │   └── app.py              # Servidor Flask principal (API + static files)
│   ├── data/
│   │   └── ficha_padrao.xlsx   # Planilha-fonte usada pelo watch_excel.ps1
│   ├── logs/                   # Logs gerados em runtime (ignorados no git)
│   ├── requirements.txt        # Dependencias Python
│   ├── scripts/
│   │   ├── start_server.bat    # Ponto de entrada principal — inicia tudo
│   │   ├── start_local_mariadb.ps1  # Sobe instancia MariaDB local
│   │   ├── stop_local_mariadb.ps1   # Para instancia MariaDB local
│   │   ├── import_dump_local.ps1    # Importa dump SQL no MariaDB local
│   │   └── watch_excel.ps1     # Monitora planilha e regenera index.html
│   └── utils/
│       ├── excel_to_html.py    # Converte aba do Excel em HTML
│       └── read_excel.py       # Leitura auxiliar de planilha
│
├── database/
│   ├── checks/
│   │   └── check_local_db.sql  # Queries rapidas de validacao do banco local
│   ├── dumps/
│   │   └── app_ferramental_dump_nodefiner.sql  # Dump sem DEFINER para carga local
│   └── scripts/
│       ├── database.sql        # Script de criacao do schema
│       └── insert_test_case.sql  # Dados de teste para desenvolvimento local
│
├── frontend/
│   ├── assets/
│   │   ├── gevernova_logo.jpg  # Logo GE usada em todas as telas
│   │   └── ficha_padrao.xlsx   # Planilha-fonte (referencia visual)
│   ├── login.html / login.css  # Tela de login
│   ├── admin.html              # Menu do administrador
│   ├── engenheiro.html         # Menu do engenheiro
│   ├── producao.html           # Pendencias da producao
│   ├── ferramenteiro.html      # Pendencias do ferramenteiro
│   ├── engenheiro_pendencias.html  # Lista completa de fichas do engenheiro
│   ├── form.html               # Formulario principal da ficha (criar/editar)
│   ├── dashboard.html          # KPIs e graficos
│   ├── planejamento.html       # Roteiro de JOB por ferramental
│   ├── planejamento_estampos.html  # Plano de manutencao de estampos
│   ├── impressao_fichas.html   # Lista de fichas para impressao
│   ├── impressao_modelo.html   # Modelo impresso da ficha individual
│   ├── consulta_vida_baixa.html  # Ferramentais com vida util < 2mm
│   ├── index.html              # Visualizacao da ficha padrao (gerado pelo Excel)
│   ├── style.css               # Estilos globais do formulario
│   ├── admin.css               # Estilos das telas de menu/painel
│   └── script.js               # Handlers globais de UI (legado)
│
└── docs/
  ├── handover.md             # Guia de transicao para novo responsavel
  └── mapeamento_backend_banco.md  # Mapeamento de campos API x banco
```

---

## 4. Perfis de usuario e fluxo de acesso

A autenticacao e **frontend-only** (sem endpoint de login no backend). O usuario seleciona seu nome na tela de login e digita um PIN. O perfil e armazenado em `sessionStorage` e cada tela verifica o perfil ao carregar, redirecionando para `login.html` se nao autorizado.

Isso significa que o login atual serve como barreira operacional de interface, nao como seguranca real de aplicacao.

### Perfis e redirecionamentos pos-login

| Perfil | Tela inicial |
|---|---|
| `admin` | `admin.html` |
| `engenheiro` | `engenheiro.html` |
| `producao` | `producao.html` |
| `ferramenteiro` | `ferramenteiro.html` |
| `planejamento` | `planejamento_estampos.html` |

### Como adicionar ou alterar usuarios

Os usuarios estao listados diretamente no `<select>` de `frontend/login.html` (linhas ~42-70). Para adicionar um usuario, inclua um `<option>` no grupo correto:

```html
<option value="perfil:Nome Completo" data-profile="perfil">Nome Completo</option>
```

Os PINs estao no bloco `<script>` da mesma tela, no objeto de credenciais. **Nao usar senhas reais — e um controle de acesso basico interno.**

---

## 5. Mapa de telas

```
login.html
    |
    +-- admin.html -------+-- engenheiro_pendencias.html
    |                     +-- ferramenteiro.html
    |                     +-- producao.html
    |                     +-- dashboard.html
    |                     +-- planejamento_estampos.html
    |                     +-- impressao_fichas.html --> impressao_modelo.html
    |                     +-- consulta_vida_baixa.html
    |                     +-- form.html (modo: impressao)
    |
    +-- engenheiro.html --+-- engenheiro_pendencias.html
    |                     +-- planejamento.html
    |                     +-- dashboard.html
    |                     +-- impressao_fichas.html --> impressao_modelo.html
    |                     +-- form.html (modo: visualizar/editar)
    |
    +-- producao.html ----+-- form.html (modo: criar)
    |
    +-- ferramenteiro.html --> form.html (modo: preencher)
    |
    +-- planejamento_estampos.html (direto para perfil planejamento)
```

### Descricao de cada tela

| Tela | Perfis com acesso | Descricao |
|---|---|---|
| `login.html` | todos | Selecao de usuario e validacao de PIN |
| `admin.html` | admin | Menu central com tiles de navegacao para todas as telas |
| `engenheiro.html` | engenheiro, admin | Menu do engenheiro com navegacao principal |
| `producao.html` | producao, admin, engenheiro | Lista de fichas pendentes da producao; abre nova ficha |
| `ferramenteiro.html` | ferramenteiro, admin, engenheiro | Fichas aguardando preenchimento do ferramenteiro |
| `engenheiro_pendencias.html` | engenheiro, admin | Tabela completa de fichas com filtros avancados |
| `form.html` | todos (modo varia por perfil) | Formulario principal da ficha com 4 secoes (ver abaixo) |
| `dashboard.html` | engenheiro, admin | KPIs e graficos de barras dos dados do banco |
| `planejamento.html` | engenheiro, admin, planejamento | Roteiro de JOB: busca ferramentais por numero de JOB |
| `planejamento_estampos.html` | todos os perfis | Tabela completa de estampos com status e vida util |
| `impressao_fichas.html` | admin, engenheiro | Lista de fichas com filtro e botao de impressao individual |
| `impressao_modelo.html` | admin, engenheiro | Layout de impressao de 2 paginas da ficha individual |
| `consulta_vida_baixa.html` | admin, engenheiro | Ferramentais com altura de vida util < 2mm |
| `index.html` | — | Preview da ficha padrao gerado pelo Excel (somente visualizacao) |

### Secoes do form.html

| ID | Nome | Preenchido por | Conteudo |
|---|---|---|---|
| `#secBasicos` | Dados Basicos | Producao | Tipo de evento, numero do ferramental, JOB, prensa, datas de entrega e emissao |
| `#secProducao` | Dados de Producao | Producao | Operador responsavel, turno, golpes por afiacao, quantidade de afiacoes |
| `#secFerramenteiro` | Dados do Ferramenteiro | Ferramenteiro | Execucao da manutencao, itens verificados, nao-conformidades, horarios |
| `#secEngenheiro` | Dados do Engenheiro | Engenheiro | Aprovacao, vida util, analytics de golpes, status final do roteiro |

---

## 6. Setup local do zero

### Pre-requisitos

- Python 3.10 ou superior
- Git configurado com acesso ao repositorio GE
- (Opcional) MariaDB local — necessario se o banco remoto nao estiver acessivel

### Passo a passo

Tempo estimado para deixar tudo rodando pela primeira vez: **10 a 20 minutos**, assumindo que as credenciais e o acesso de rede ja estao liberados.

**1. Clonar o repositorio**

```powershell
cd C:\Users\SEU_USUARIO\Documents
git clone https://github.apps.gevernova.net/212806893/Tooling_Production.git
cd Tooling_Production
code .
```

**2. Criar o ambiente virtual Python**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

> Se o PowerShell bloquear a execucao de scripts:
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
> ```

**3. Instalar dependencias**

```powershell
pip install -r backend\requirements.txt
```

**4. Configurar o arquivo `.env`**

Copie o template e preencha com as credenciais reais:

```powershell
Copy-Item .env.example .env
notepad .env
```

Veja a secao [Variaveis de ambiente](#7-variaveis-de-ambiente) para detalhes.

**5. Iniciar o servidor**

```powershell
backend\scripts\start_server.bat
```

**6. Acessar no navegador**

```
http://localhost:5000
```

O servidor redireciona automaticamente para a tela de login.

---

## 7. Variaveis de ambiente

O arquivo `.env` na raiz do projeto configura a conexao com o banco. **Nunca versionar o `.env` real.**

### O que cada variavel controla

```env
DB_HOST=seu-host-mariadb.local   # Host do banco MariaDB da GE
DB_PORT=3308                        # Porta (diferente da padrao 3306)
DB_NAME=app_ferramental          # Nome do banco
DB_USER=APP_DB_USER                     # Usuario do banco
DB_PASSWORD=COLOQUE_SUA_SENHA_AQUI  # Senha — obter com o gestor do banco
DB_SSL_MODE=DISABLED                # SSL desabilitado na rede interna GE
```

| Variavel | Funcao | Observacao |
|---|---|---|
| `DB_HOST` | Define para qual banco a aplicacao aponta | Troque para `127.0.0.1` quando usar banco local |
| `DB_PORT` | Porta TCP da conexao | Remoto usa `3308`; local usa `3307` |
| `DB_NAME` | Schema principal da aplicacao | Hoje o projeto espera `app_ferramental` |
| `DB_USER` | Usuario da conexao | No remoto e `APP_DB_USER`; no local pode ser `root` |
| `DB_PASSWORD` | Senha da conexao | Deve vir do responsavel tecnico ou ser vazia no local |
| `DB_SSL_MODE` | Modo SSL do conector | Hoje o ambiente interno opera com `DISABLED` |

> **Para banco local (MariaDB local de desenvolvimento):**
> ```env
> DB_HOST=127.0.0.1
> DB_PORT=3307
> DB_NAME=app_ferramental
> DB_USER=root
> DB_PASSWORD=
> ```
> Use `backend\scripts\start_local_mariadb.ps1` para subir a instancia local.

---

## 8. Como iniciar o servidor

### Forma principal (recomendada)

```bat
backend\scripts\start_server.bat
```

Este script automaticamente:
- Usa o Python da `.venv` na raiz
- Tenta iniciar o MariaDB local se disponivel
- Sobe a API Flask em `0.0.0.0:5000`
- Registra crashes em `backend/logs/server_crash.log`

### Ordem recomendada para subir sem erro

1. Confirmar `.venv` pronta.
2. Confirmar `.env` preenchido.
3. Decidir se vai usar banco GE ou banco local.
4. Rodar `backend/scripts/start_server.bat`.
5. Testar `/health` antes de abrir varias telas.

### Forma direta (debug)

```powershell
.\.venv\Scripts\Activate.ps1
python backend\api\app.py
```

### Verificar se esta funcionando

| URL | Esperado |
|---|---|
| `http://localhost:5000` | Redireciona para `login.html` |
| `http://localhost:5000/health` | JSON `{"status": "ok", "db": "conectado"}` |
| `http://localhost:5000/frontend/login.html` | Tela de login |

---

## 9. Endpoints de API

Todos os endpoints retornam JSON. O servidor tambem serve os arquivos estaticos do `frontend/`.

### Arquivos estaticos

| Metodo | Rota | Descricao |
|---|---|---|
| `GET` | `/` | Redireciona para `frontend/login.html` |
| `GET` | `/<path>` | Serve qualquer arquivo da pasta `frontend/` |

### Fichas

| Metodo | Rota | Descricao |
|---|---|---|
| `GET` | `/api/fichas` | Lista todas as fichas com suporte a filtros |
| `POST` | `/api/fichas` | Cria uma nova ficha |
| `GET` | `/api/ficha/<id>` | Retorna o detalhe de uma ficha pelo ID |
| `PUT` | `/api/ficha/<id>` | Atualiza uma ficha existente |
| `DELETE` | `/api/ficha/<id>` | Remove uma ficha |

#### Filtros disponiveis em `GET /api/fichas`

Parametros de query string (todos opcionais):

| Parametro | Tipo | Exemplo | Descricao |
|---|---|---|---|
| `job` | string | `?job=2280091` | Filtra pelo numero da JOB |
| `ferramental` | string | `?ferramental=F286` | Filtro parcial no numero do ferramental |
| `prensa` | string | `?prensa=PX59` | Filtra pela prensa |
| `status` | string | `?status=pendente` | Status da ficha |
| `status_ferramental` | string | `?status_ferramental=ativo` | Status do ferramental |
| `data_inicio` | date | `?data_inicio=2026-01-01` | Data minima de criacao |
| `data_fim` | date | `?data_fim=2026-03-31` | Data maxima de criacao |

#### Exemplo de resposta `GET /api/fichas`

```json
[
  {
    "id": 44684,
    "job_principal": "2280091",
    "numero_ferramenta": "F2860791-0010",
    "descricao_ferramenta": "RANHURADOR WuHu",
    "prensa": "PX59",
    "status_roteiro": "pendente",
    "data_entrega_evento": "2026-02-11",
    "vida_util_atual_mm": 3.2
  }
]
```

### Outros endpoints

| Metodo | Rota | Descricao |
|---|---|---|
| `GET` | `/api/estatisticas` | KPIs para o dashboard (totais por status, por prensa etc.) |
| `GET` | `/api/exportar-excel` | Baixa todas as fichas como arquivo `.xlsx` |
| `GET` | `/health` | Health check da API e conectividade com o banco |

---

## 10. Banco de dados

### Conexao

O banco e um **MariaDB** hospedado nos servidores da GE:

- **Host:** `seu-host-mariadb.local`
- **Porta:** `3308` (nao e a porta padrao — atencao ao configurar clientes SQL)
- **Banco:** `app_ferramental`
- **Usuario:** `APP_DB_USER`
- **Senha:** obtida com o gestor tecnico / DBA da unidade

### O que considerar como fonte de verdade

| Artefato | Como tratar |
|---|---|
| Banco remoto GE | Fonte de dados real de operacao |
| `database/dumps/app_ferramental_dump_nodefiner.sql` | Melhor base para reproduzir dados localmente |
| `database/scripts/database.sql` | Referencia de schema e apoio para entendimento, nao substitui o banco real |
| `database/checks/check_local_db.sql` | Apoio para validacao e diagnostico |

### Banco local para desenvolvimento

Para trabalhar sem acesso ao servidor de producao:

1. Subir instancia local:
   ```powershell
   backend\scripts\start_local_mariadb.ps1
   ```

2. Importar o dump:
   ```powershell
   backend\scripts\import_dump_local.ps1
   ```
   O dump esta em `database\dumps\app_ferramental_dump_nodefiner.sql`.
   O script ja remove os `DEFINER` automaticamente para que o import funcione localmente.

3. Atualizar o `.env` para apontar para `127.0.0.1:3307`.

4. Para parar a instancia local:
   ```powershell
   backend\scripts\stop_local_mariadb.ps1
   ```

### Queries de validacao rapida

```sql
-- Ver quantidade de fichas por status
-- Arquivo: database/checks/check_local_db.sql
```

Execute via qualquer cliente SQL (DBeaver, HeidiSQL, linha de comando) ou pela query direta no terminal.

### Schema

O schema de referencia esta em `database\scripts\database.sql`. A tabela principal em producao e `banco_dados` no schema `app_ferramental`.

| Coluna | Tipo | Descricao |
|---|---|---|
| `numero_ficha` | INT | Identificador da ficha (PK) |
| `job` | VARCHAR | Numero da JOB principal |
| `job_envolvidas` | VARCHAR | JOBs secundarias relacionadas |
| `elemento_pep` | VARCHAR | Elemento PEP do projeto |
| `tipo_lamina_processo` | VARCHAR | Tipo de lamina / processo |
| `desenhos` | VARCHAR | Referencias de desenhos |
| `numero_ferramental` | VARCHAR | Numero do ferramental (ex: F2860791-0010) |
| `descricao_ferramental` | VARCHAR | Descricao do ferramental |
| `ficha_processo` | VARCHAR | Referencia da ficha de processo |
| `numero_desenho_ferramental` | VARCHAR | Numero do desenho do ferramental |
| `pt` | VARCHAR | PT associado |
| `prensa` | VARCHAR | Prensa utilizada (ex: PX59) |
| `status_roteiro` | VARCHAR | Status atual da ficha no roteiro |
| `status_ferramental` | VARCHAR | Status do ferramental (ativo, inativo etc.) |
| `status_frnr` | VARCHAR | Status FRNR |
| `status_entrega_ferramental` | VARCHAR | Status de entrega do ferramental |
| `data_entrega_evento` | DATE | Data prevista de entrega |
| `data_emissao_requisicao` | DATE | Data de emissao da requisicao |
| `data_termino_producao` | DATE | Data de termino na producao |
| `data_termino_avaliacao_kit` | DATE | Data de termino da avaliacao do kit |
| `data_envio_ferramentaria` | DATE | Data de envio para a ferramentaria |
| `operador_responsavel` | VARCHAR | Operador que abriu a ficha |
| `ferramenteiro_responsavel` | VARCHAR | Ferramenteiro que executou a manutencao |
| `horario_inicio_ferramenteiro` | TIME | Hora de inicio da manutencao |
| `horario_termino_ferramenteiro` | TIME | Hora de termino da manutencao |
| `qtde_afiacao` | INT | Quantidade de afiacoes realizadas |
| `qtde_total_afiacao` | INT | Total acumulado de afiacoes |
| `qtde_total_golpes_ferramenta` | INT | Total de golpes do ferramental |
| `qtde_uso_ferramenta` | INT | Quantidade de usos do ferramental |
| `total_golpes_final` | INT | Total de golpes ao final |
| `total_golpes_sem_afiar` | INT | Golpes acumulados sem afiacao |
| `1_golpes_afiacao` .. `5_golpes_afiacao` | INT | Golpes por ciclo de afiacao (ate 5) |
| `vida_util_ferramenta` | DECIMAL | Vida util restante em mm |
| `observacao_ultima_manutencao_ferramenta` | TEXT | Observacoes da ultima manutencao |
| `observacao_rnc` | TEXT | Observacoes de nao-conformidade |
| `numero_rnc` | VARCHAR | Numero do RNC |
| `numero_plan_interno` | VARCHAR | Numero do plano interno |
| `numero_po` | VARCHAR | Numero da PO |

---

## 11. Scripts operacionais

Todos os scripts ficam em `backend\scripts\`.

| Script | Como executar | O que faz |
|---|---|---|
| `start_server.bat` | `backend\scripts\start_server.bat` | Inicia a API Flask usando o Python do `.venv`. Ponto principal de entrada. |
| `start_local_mariadb.ps1` | `.\backend\scripts\start_local_mariadb.ps1` | Sobe a instancia MariaDB portable local (porta 3307). |
| `stop_local_mariadb.ps1` | `.\backend\scripts\stop_local_mariadb.ps1` | Para graciosamente a instancia MariaDB local. |
| `import_dump_local.ps1` | `.\backend\scripts\import_dump_local.ps1` | Importa o dump SQL no banco local, removendo definers para evitar erros de permissao. |
| `watch_excel.ps1` | `.\backend\scripts\watch_excel.ps1` | Monitora a planilha-fonte (definida internamente no script) e regenera `frontend\index.html` sempre que o arquivo e salvo. Utilitario de desenvolvimento. |

### Utilitarios Python

| Script | Como executar | O que faz |
|---|---|---|
| `backend\utils\excel_to_html.py` | `python backend\utils\excel_to_html.py --input caminho.xlsx` | Converte uma aba do Excel em HTML preservando o estilo visual. Usado pelo `watch_excel.ps1`. |
| `backend\utils\read_excel.py` | `python backend\utils\read_excel.py` | Script auxiliar de leitura e inspecao de planilha. |

---

## 12. Fluxo Git

O repositorio esta hospedado no GitHub interno da GE:
`https://github.apps.gevernova.net/212806893/Tooling_Production`

### Inicio do dia (sincronizar antes de comecar)

```powershell
git checkout main
git pull origin main
git checkout -b feature/nome-da-feature
```

### Durante o desenvolvimento (commits frequentes)

```powershell
git add .
git commit -m "feat: descricao objetiva da alteracao"
```

### Final do dia (publicar)

```powershell
git push -u origin feature/nome-da-feature
```

Depois abrir um **Pull Request** no GitHub GE de `feature/nome-da-feature` para `main`.

### Trabalho direto na main (somente para hotfixes urgentes)

```powershell
git pull origin main
git add .
git commit -m "fix: descricao do hotfix"
git push origin main
```

### Resolver conflito de merge

```powershell
git status                          # Ver arquivos em conflito
# Editar os arquivos marcados com <<< >>> ===
git add .
git commit -m "merge: resolve conflicts"
```

---

## 13. Convencoes de commit

| Prefixo | Quando usar |
|---|---|
| `feat:` | Nova funcionalidade ou tela |
| `fix:` | Correcao de bug |
| `refactor:` | Reorganizacao de codigo sem mudar comportamento |
| `chore:` | Tarefa tecnica (dependencias, scripts, gitignore) |
| `docs:` | Atualizacao de documentacao |
| `style:` | Ajuste visual / CSS |
| `db:` | Alteracao de schema ou dados do banco |

**Exemplos bons:**
```
feat: adicionar filtro por prensa na tela de planejamento
fix: corrigir loop infinito no carregamento do dashboard
docs: atualizar README com instrucoes de banco local
```

---

## 14. Troubleshooting

### Servidor nao sobe

```powershell
# 1. Verificar se o .venv existe e esta ativo
.\.venv\Scripts\Activate.ps1

# 2. Reinstalar dependencias
pip install -r backend\requirements.txt

# 3. Rodar diretamente para ver o erro completo
python backend\api\app.py
```

### Erro de conexao com o banco

- Confirmar que o `.env` existe e tem os valores corretos.
- Testar acesso a rede: `Test-NetConnection seu-host-mariadb.local -Port 3308`
- Se a rede GE nao estiver acessivel (ex: home office sem VPN), usar o banco local.
- Verificar o health check: `http://localhost:5000/health`
- Se estiver usando banco local, confirmar se a porta `3307` esta ativa antes de culpar o Flask.

### Tela fica em branco ou redireciona para login em loop

- Abrir o Console do navegador (F12 > Console) e ver o erro JavaScript.
- Geralmente indica que `sessionStorage.getItem('perfil_ativo')` esta nulo — limpar o sessionStorage:
  ```javascript
  // No console do navegador:
  sessionStorage.clear();
  location.reload();
  ```

### Erro de autenticacao Git

```powershell
git config --global --unset credential.helper
git config --global credential.helper manager
# Tentar o push novamente — uma janela de login sera exibida
```

### Politica de execucao do PowerShell bloqueia scripts

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### MariaDB local nao sobe

- Verificar se a pasta `.mariadb_local/` existe na raiz.
- Abrir o `start_local_mariadb.ps1` e confirmar os caminhos do executavel.
- Checar se a porta 3307 nao esta em uso: `netstat -an | findstr 3307`

### O sistema subiu, mas os dados parecem incoerentes

- Confirme primeiro se voce esta olhando para banco remoto ou local.
- Verifique se o dump local usado esta atualizado.
- Compare uma ficha especifica direto no banco antes de concluir que o frontend esta errado.
- Se o erro aparecer so em uma tela, mapeie a consulta via secao 9 e inspecione o endpoint correspondente.

---

## 15. Roadmap e proximos passos

Itens tecnicos identificados para quem der continuidade ao projeto:

| Prioridade | Item | Detalhes |
|---|---|---|
| Alta | **Autenticacao real** | O login atual e frontend-only. Implementar validacao de PIN/usuario no backend com sessao ou JWT. |
| Alta | **Modularizar `app.py`** | O arquivo tem ~970 linhas concentrando rotas, queries, serializacao e logica de negocio. Separar em `routes/`, `services/` e `db/`. |
| Media | **Ambiente de staging** | Criar um `.env` separado apontando para banco de homologacao antes de subir producao. |
| Media | **Testes automatizados** | Nao ha nenhum teste. Comecar com testes de endpoint (`pytest` + `flask test client`). |
| Baixa | **Cache de API** | O endpoint `GET /api/fichas` ja tem cache em memoria. Avaliar TTL e consistencia ao salvar fichas. |
| Baixa | **Logs estruturados** | Substituir prints por `logging` com nivel e formato JSON para facilitar monitoramento. |

---

## Contato e repositorio

- **Repositorio:** `https://github.apps.gevernova.net/212806893/Tooling_Production`
- **Responsavel anterior:** Gustavo Nascimento (AME Resp.)
- **Documentacao tecnica adicional:** [`docs/handover.md`](docs/handover.md)

