"""
Arquivo: backend/api/app.py
Descricao: servidor Flask principal do sistema de fichas de manutencao.

Responsabilidades principais:
- Expor endpoints de API REST para consulta, filtro e exportacao de fichas.
- Servir arquivos estaticos do frontend (HTML, CSS, JS) com suporte a auto-reload local.
- Centralizar acesso ao banco MySQL/MariaDB com serializacao consistente de tipos Python.
- Aplicar compressao gzip em respostas JSON elegiveis para reduzir trafego de rede.
- Manter cache em memoria para consultas de alta frequencia com TTL configuravel.

Observacao sobre arquitetura:
# Notas de manutencao:
# - Objetivo: manter comportamento funcional sem quebrar compatibilidade.
# - Cuidado: validar efeitos em fluxo web, banco e exportacoes.
# - Ao alterar: preferir mudancas pequenas e validacao local.

- Este modulo concentra muitas responsabilidades por razoes historicas (legado).
- A evolucao recomendada e modularizar em camadas separadas: db/ + services/ + routes/.
"""

# ---------------------------------------------------------------------------
# Imports da biblioteca padrao do Python
# ---------------------------------------------------------------------------
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
import json
import os
import subprocess
import gzip
import threading
import time

# ---------------------------------------------------------------------------
# Imports de bibliotecas de terceiros
# ---------------------------------------------------------------------------
from flask import Flask, jsonify, make_response, request, send_file, send_from_directory
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Conector MySQL/MariaDB — importado com tratamento de ausencia.
# Quando mysql-connector-python nao esta instalado, a API usa o cliente CLI
# do MariaDB como fallback (via subprocess). Ver funcao _consultar_todos_cli().
try:
    import mysql.connector
    from mysql.connector import Error as MySQLError
except ImportError:
    mysql = None
    MySQLError = Exception

# python-dotenv para carregar variaveis de ambiente do arquivo .env.
# Quando nao disponivel, o .env e lido manualmente linha por linha.
try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

# ---------------------------------------------------------------------------
# Caminhos absolutos do projeto — derivados do local deste arquivo.
# Estrutura esperada: RAIZ_PROJETO/backend/api/app.py
# ---------------------------------------------------------------------------
DIRETORIO_ATUAL   = os.path.dirname(os.path.abspath(__file__))     # .../backend/api
DIRETORIO_BACKEND = os.path.dirname(DIRETORIO_ATUAL)               # .../backend
RAIZ_PROJETO      = os.path.dirname(DIRETORIO_BACKEND)             # raiz do repositorio
DIRETORIO_FRONTEND = os.path.join(RAIZ_PROJETO, "frontend")        # pasta com HTML/CSS/JS
CAMINHO_DOTENV    = os.path.join(RAIZ_PROJETO, ".env")             # arquivo de configuracao local

# ---------------------------------------------------------------------------
# Carregamento das variaveis de ambiente a partir do .env
# ---------------------------------------------------------------------------
if load_dotenv:
    # Metodo ideal: delega para a biblioteca python-dotenv
    load_dotenv(CAMINHO_DOTENV)
else:
    # Fallback manual: le o .env e injeta as variaveis em os.environ
    if os.path.exists(CAMINHO_DOTENV):
        with open(CAMINHO_DOTENV, "r", encoding="utf-8") as arquivo_env:
            for linha_bruta in arquivo_env:
                linha = linha_bruta.strip()
                # Pula linhas vazias e comentarios
                if not linha or linha.startswith("#") or "=" not in linha:
                    continue
                chave, valor = linha.split("=", 1)
                chave = chave.strip()
                valor = valor.strip().strip('"').strip("'")
                # Nao sobrescreve variaveis ja definidas no ambiente do sistema operacional
                if chave and chave not in os.environ:
                    os.environ[chave] = valor

# ---------------------------------------------------------------------------
# Inicializacao da aplicacao Flask e CORS
# ---------------------------------------------------------------------------
app = Flask(__name__, static_folder=None)
CORS(app)  # Habilita cross-origin — necessario para dev local com portas diferentes

# Tamanho maximo (em bytes) de uma resposta JSON para aplicar compressao gzip.
# Respostas acima desse limite sao enviadas sem compressao para evitar pico de memoria.
LIMITE_BYTES_GZIP = int(os.getenv("LIMITE_BYTES_GZIP", str(1024 * 1024)))

# ---------------------------------------------------------------------------
# Cache em memoria para respostas de API
# Objetivo: evitar consultas repetidas ao banco para filtros identicos.
# Implementacao: dicionario com TTL por entrada + lock de thread para seguranca.
# ---------------------------------------------------------------------------
_TRAVA_CACHE = threading.Lock()  # Mutex — garante acesso atomico ao dicionario de cache
_CACHE: dict = {}                # chave → {dados, expira_em, ultimo_acesso}
_CACHE_CAPACIDADE = 128          # Numero maximo de entradas; entradas antigas sao evicadas

SCRIPT_AUTO_RELOAD_DEV = """
<script>
(function () {
    // Executa somente em ambiente local — producao nao carrega esse script
    if (window.location.hostname !== "localhost" && window.location.hostname !== "127.0.0.1") {
        return;
    }
    let ultimoTimestamp = 0;

    async function verificarRecarga() {
        try {
            const resposta = await fetch("/__dev_timestamp?ts=" + Date.now(), { cache: "no-store" });
            if (!resposta.ok) return;
            const dados = await resposta.json();
            const timestampAtual = Number(dados.ts || 0);
            if (!timestampAtual) return;
            // Se o timestamp subiu desde a ultima checagem, um arquivo foi modificado
            if (ultimoTimestamp && timestampAtual > ultimoTimestamp) {
                window.location.reload();
                return;
            }
            ultimoTimestamp = timestampAtual;
        } catch (_erro) {
            // Silencioso — falha de rede nao deve travar a experiencia de dev
        }
    }

    verificarRecarga();
    setInterval(verificarRecarga, 1000);  // Intervalo de polling: 1 segundo
})();
</script>
"""


def _serve_html_with_dev_reload(caminho_arquivo: str):
    """Le um arquivo HTML e injeta o snippet de auto-reload para uso local.

    O snippet so e ativado em localhost/127.0.0.1, portanto e seguro injetar
    em todos os HTMLs sem risco de impactar producao.
    """
    # Tenta UTF-8 primeiro; fallback latin-1 para arquivos legados
    try:
        with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
            conteudo_html = arquivo.read()
    except UnicodeDecodeError:
        with open(caminho_arquivo, "r", encoding="latin-1") as arquivo:
            conteudo_html = arquivo.read()

    # Injeta antes do </body> para nao bloquear renderizacao
    if "</body>" in conteudo_html:
        conteudo_html = conteudo_html.replace("</body>", SCRIPT_AUTO_RELOAD_DEV + "\n</body>", 1)
    else:
        conteudo_html = conteudo_html + SCRIPT_AUTO_RELOAD_DEV

    resposta = make_response(conteudo_html)
    # Cabecalhos no-cache: browser sempre busca a versao mais recente
    resposta.headers["Content-Type"] = "text/html; charset=utf-8"
    resposta.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resposta.headers["Pragma"] = "no-cache"
    resposta.headers["Expires"] = "0"
    return resposta


def _cache_get(chave):
    """Retorna dados armazenados para a chave se ainda validos; caso contrario None.

    Entradas expiradas sao removidas no momento da consulta (lazy expiry).
    """
    agora = time.monotonic()
    with _TRAVA_CACHE:
        entrada = _CACHE.get(chave)
        if not entrada:
            return None
        # Entrada expirou — remove e informa que nao ha cache valido
        if entrada["expires_at"] <= agora:
            _CACHE.pop(chave, None)
            return None
        # Atualiza horario de acesso para que a evicao LRU funcione corretamente
        entrada["last_hit"] = agora
        return entrada["payload"]


def _cache_set(chave, dados, ttl_segundos):
    """Armazena dados com TTL e evica entradas antigas quando o cache esta cheio.

    Estrategia de evicao:
    1. Remove entradas expiradas (custo zero de reputacao).
    2. Se ainda acima da capacidade, remove as menos acessadas recentemente (LRU).
    """
    agora = time.monotonic()
    expira_em = agora + max(1.0, float(ttl_segundos))
    with _TRAVA_CACHE:
        # Armazena a entrada com metadados de controle de TTL e acesso
        _CACHE[chave] = {
            "payload": dados,
            "expires_at": expira_em,
            "last_hit": agora,
        }

        # Capacidade OK — nada a evictar
        if len(_CACHE) <= _CACHE_CAPACIDADE:
            return

        # Passo 1: remove entradas cujo TTL ja expirou
        chaves_expiradas = [k for k, v in _CACHE.items() if v["expires_at"] <= agora]
        for chave_expirada in chaves_expiradas:
            _CACHE.pop(chave_expirada, None)

        if len(_CACHE) <= _CACHE_CAPACIDADE:
            return

        # Passo 2: evica as entradas acessadas ha mais tempo (LRU)
        chaves_por_acesso = sorted(_CACHE.keys(), key=lambda k: _CACHE[k]["last_hit"])
        excesso = len(_CACHE) - _CACHE_CAPACIDADE
        for chave_antiga in chaves_por_acesso[:excesso]:
            _CACHE.pop(chave_antiga, None)


def _db_config():
    """Monta o dicionario de configuracao de conexao a partir das variaveis de ambiente.

    Variaveis esperadas no .env: DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME.
    DB_SSL_MODE pode ser "DISABLED" ou "REQUIRED" para controle de SSL.
    """
    configuracao = {
        "host":       os.getenv("DB_HOST", "localhost"),
        "port":       int(os.getenv("DB_PORT", "3306")),
        "user":       os.getenv("DB_USER", ""),
        "password":   os.getenv("DB_PASSWORD", ""),
        "database":   os.getenv("DB_NAME", ""),
        "charset":    "utf8mb4",   # Suporte a acentos e caracteres especiais
        "use_unicode": True,
        "autocommit": False,       # Transacoes confirmadas explicitamente
    }

    modo_ssl = os.getenv("DB_SSL_MODE", "").strip().upper()
    if modo_ssl == "DISABLED":
        configuracao["ssl_disabled"] = True
    if modo_ssl == "REQUIRED":
        configuracao["ssl_disabled"] = False

    return configuracao


def _connect_db():
    """Cria e retorna uma conexao com o banco usando mysql-connector.

    Levanta RuntimeError se o conector nao estiver instalado ou se as
    variaveis de ambiente obrigatorias estiverem ausentes no .env.
    """
    if mysql is None:
        raise RuntimeError("Dependencia mysql-connector-python nao instalada.")

    conf = _db_config()
    # Verifica se as variaveis essenciais foram configuradas no .env
    ausentes = [var for var in ("DB_USER", "DB_PASSWORD", "DB_NAME") if not os.getenv(var)]
    if ausentes:
        raise RuntimeError("Variaveis ausentes no .env: " + ", ".join(ausentes))

    return mysql.connector.connect(**conf)


def _escape_sql_value(valor):
    """Escapa um valor Python para interpolacao segura em SQL literal.

    Usado apenas no modo CLI (sem mysql-connector). Com o conector nativo
    os parametros sao passados de forma segura via placeholders %s.
    """
    if valor is None:
        return "NULL"
    if isinstance(valor, bool):
        return "1" if valor else "0"
    if isinstance(valor, (int, float, Decimal)):
        return str(valor)
    texto = str(valor)
    # Escapa barras inversas e aspas simples para prevenir injecao SQL
    texto = texto.replace("\\", "\\\\").replace("'", "\\'")
    return f"'{texto}'"


def _interpolate_query(consulta: str, parametros=None) -> str:
    """Substitui os placeholders %s da consulta pelos valores escapados.

    Usado apenas no caminho CLI. Levanta RuntimeError se a quantidade de
    parametros nao corresponder ao numero de placeholders.
    """
    if not parametros:
        return consulta
    partes = consulta.split("%s")
    if len(partes) - 1 != len(parametros):
        raise RuntimeError(
            f"Quantidade de parametros ({len(parametros)}) nao confere "
            f"com os placeholders na consulta ({len(partes) - 1})."
        )
    resultado = [partes[0]]
    for i, valor in enumerate(parametros):
        resultado.append(_escape_sql_value(valor))
        resultado.append(partes[i + 1])
    return "".join(resultado)


def _query_all_cli(consulta: str, parametros=None) -> list:
    """Executa SELECT via cliente CLI do MariaDB e retorna lista de dicionarios.

    Usado quando mysql-connector nao esta disponivel.
    A saida TSV do cliente e parseada manualmente em registros Python.
    """
    conf = _db_config()
    # Verifica variaveis de conexao necessarias para o modo CLI
    ausentes = [v for v in ("DB_HOST", "DB_PORT", "DB_USER", "DB_PASSWORD", "DB_NAME") if not os.getenv(v)]
    if ausentes:
        raise RuntimeError("Variaveis de ambiente ausentes: " + ", ".join(ausentes))

    exe_cliente = os.getenv("DB_CLIENT_EXE", r"C:\Program Files\MariaDB 12.2\bin\mariadb.exe")
    if not os.path.exists(exe_cliente):
        raise RuntimeError(f"Cliente MariaDB nao encontrado em: {exe_cliente}")

    # Interpola os parametros no SQL (modo CLI nao suporta placeholders nativos)
    sql = _interpolate_query(consulta, parametros)
    argumentos = [
        exe_cliente,
        "--skip-ssl",
        "--batch",          # Saida em formato TSV, sem decoracao de tabela
        "-h", str(conf["host"]),
        "-P", str(conf["port"]),
        "-u", str(conf["user"]),
        f"-p{conf['password']}",
        str(conf["database"]),
        "-e", sql,
    ]

    # Executa o cliente de linha de comando e captura a saida
    processo = subprocess.run(
        argumentos,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if processo.returncode != 0:
        mensagem_erro = (processo.stderr or "").strip()
        raise RuntimeError(mensagem_erro or "Falha ao executar consulta via cliente MariaDB.")

    # Parseia saida TSV: primeira linha = cabecalhos, demais = dados
    linhas = [l for l in (processo.stdout or "").splitlines() if l.strip() != ""]
    if not linhas:
        return []

    cabecalhos = linhas[0].split("\t")
    registros = []
    for linha in linhas[1:]:
        colunas = linha.split("\t")
        registro = {}
        for i, coluna in enumerate(cabecalhos):
            registro[coluna] = colunas[i] if i < len(colunas) else None
        registros.append(_serialize_row(registro))
    return registros


def _serialize_value(valor):
    """Converte um valor retornado pelo banco em tipo compativel com JSON.

    - str: passa pela correcao de mojibake (encoding duplo UTF-8/latin1)
    - datetime/date: formato ISO 8601
    - Decimal: float
    - Outros: retorna sem alteracao
    """
    if isinstance(valor, str):
        return _fix_mojibake(valor)
    if isinstance(valor, (datetime, date)):
        return valor.isoformat()
    if isinstance(valor, Decimal):
        return float(valor)
    return valor


def _fix_mojibake(texto: str) -> str:
    """Corrige texto com dupla codificacao UTF-8/latin1 (mojibake).

    Problema comum em bancos migrados de MySQL com charset mal configurado.
    Exemplo: "ESPAÃƒâ€¡ADORES" deve ser "ESPACADORES".
    Tenta ate 3 rodadas de encode(latin1).decode(utf-8) ate o texto estabilizar.
    """
    if not isinstance(texto, str):
        return texto
    # Verificacao rapida — maioria dos textos nao precisa de correcao
    if "Ãƒ" not in texto and "Ã‚" not in texto and "Ã¢" not in texto:
        return texto

    corrigido = texto
    for _ in range(3):
        try:
            candidato = corrigido.encode("latin1", errors="strict").decode("utf-8", errors="strict")
        except Exception:
            break  # Nao e mais possivel corrigir — retorna estado atual
        if candidato == corrigido:
            break  # Texto estabilizou — nao ha mais correcoes
        corrigido = candidato
        if "Ãƒ" not in corrigido and "Ã‚" not in corrigido and "Ã¢" not in corrigido:
            break  # Texto limpo — encerra a correcao
    return corrigido


def _serialize_row(registro: dict) -> dict:
    """Aplica _serialize_value em todos os campos de um registro do banco."""
    return {chave: _serialize_value(valor) for chave, valor in registro.items()}


def _normalize_status(status_roteiro: str) -> str:
    """Converte o texto bruto de status_roteiro do banco para uma chave de status do frontend.

    Mapeamento:
    - None / "na" / vazio → "pendente"
    - "aprovado" / "conclu*" → "concluida"
    - "andamento" / "execu*" → "em_andamento"
    - "roteiro" → "pendente"
    - Qualquer outro → "em_andamento" (padrao conservador)
    """
    valor = (status_roteiro or "").strip().lower()
    if not valor or valor == "na":
        return "pendente"
    if "aprovado" in valor or "conclu" in valor:
        return "concluida"
    if "andamento" in valor or "execu" in valor:
        return "em_andamento"
    if "roteiro" in valor:
        return "pendente"
    return "em_andamento"


def _map_banco_dados_to_api(registro: dict) -> dict:
    """Transforma um registro da tabela banco_dados no contrato de payload da API.

    Campos tecnicos do banco (ex: tipo_lamina_processo) recebem aliases semanticos
    (ex: equipamento, tag) alem de manterem os nomes originais para compatibilidade.
    """
    # Calcula o status normalizado a partir do campo bruto do banco
    status = _normalize_status(registro.get("status_roteiro"))
    return {
        "id": registro.get("numero_ficha"),
        "numero_ficha": registro.get("numero_ficha"),
        "equipamento": registro.get("tipo_lamina_processo"),
        "tag": registro.get("numero_ferramental"),
        "tipo_manutencao": registro.get("tipo_lamina_processo"),
        "tipo_lamina_processo": registro.get("tipo_lamina_processo"),
        "data_execucao": _serialize_value(registro.get("data_termino_producao")),
        "responsavel": registro.get("operador_responsavel"),
        "area": registro.get("pt"),
        "descricao": registro.get("descricao_ferramental"),
        "observacoes": registro.get("observacao_ultima_manutencao_ferramenta") or registro.get("observacao_rnc"),
        "tempo_execucao": registro.get("vida_util_ferramenta"),
        "status": status,
        "data_criacao": _serialize_value(registro.get("data_emissao_requisicao")),
        "data_atualizacao": _serialize_value(registro.get("data_termino_producao")),
        # Campos esperados pelo frontend atual:
        "descricao_ferramenta": registro.get("descricao_ferramental"),
        "codigo_ferramenta": registro.get("numero_ferramental"),
        "documento": registro.get("numero_plan_interno") or registro.get("numero_po"),
        "data_cadastro": _serialize_value(registro.get("data_emissao_requisicao")),
        "responsavel_abertura": registro.get("operador_responsavel"),
        "ferramenteiro": registro.get("ferramenteiro_responsavel"),
        "qtde_afiacao": registro.get("qtde_afiacao"),
        "data_termino": _serialize_value(registro.get("data_termino_producao")),
        "descricao_trabalhos": registro.get("tipo_lamina_processo"),
        "job": registro.get("job"),
        "job_envolvidas": registro.get("job_envolvidas"),
        "elemento_pep": registro.get("elemento_pep"),
        "desenhos": registro.get("desenhos"),
        "pt": registro.get("pt"),
        "numero_ferramental": registro.get("numero_ferramental"),
        "ficha_processo": registro.get("ficha_processo"),
        "numero_desenho_ferramental": registro.get("numero_desenho_ferramental"),
        "status_roteiro": registro.get("status_roteiro"),
        "status_frnr": registro.get("status_frnr"),
        "status_ferramental": registro.get("status_ferramental"),
        "status_ferramental_lob": registro.get("status_ferramental"),
        "status_entrega_ferramental": registro.get("status_entrega_ferramental"),
        "data_entrega_evento": _serialize_value(registro.get("data_entrega_evento")),
        "data_emissao_requisicao": _serialize_value(registro.get("data_emissao_requisicao")),
        "data_termino_avaliacao_kit": _serialize_value(registro.get("data_termino_avaliacao_kit")),
        "data_termino_producao": _serialize_value(registro.get("data_termino_producao")),
        "prensa": registro.get("prensa"),
        "numero_ferramenta": registro.get("numero_ferramental"),
        "numero_desenho_ferramenta": registro.get("numero_desenho_ferramental"),
        "numero_plan_interno": registro.get("numero_plan_interno"),
        "numero_po": registro.get("numero_po"),
        "numero_rnc": registro.get("numero_rnc"),
        "observacao_rnc": registro.get("observacao_rnc"),
        "total_golpes_final": registro.get("total_golpes_final"),
        "total_golpes_sem_afiar": registro.get("total_golpes_sem_afiar"),
        "qtde_total_golpes_ferramenta": registro.get("qtde_total_golpes_ferramenta"),
        "golpes_afiacao_1": registro.get("1_golpes_afiacao"),
        "golpes_afiacao_2": registro.get("2_golpes_afiacao"),
        "golpes_afiacao_3": registro.get("3_golpes_afiacao"),
        "golpes_afiacao_4": registro.get("4_golpes_afiacao"),
        "golpes_afiacao_5": registro.get("5_golpes_afiacao"),
        "qtde_total_afiacao": registro.get("qtde_total_afiacao"),
        "qtde_uso_ferramenta": registro.get("qtde_uso_ferramenta"),
        "vida_util_ferramenta": registro.get("vida_util_ferramenta"),
        "data_envio_ferramentaria": _serialize_value(registro.get("data_envio_ferramentaria")),
        "horario_inicio_ferramenteiro": _serialize_value(registro.get("horario_inicio_ferramenteiro")),
        "horario_termino_ferramenteiro": _serialize_value(registro.get("horario_termino_ferramenteiro")),
    }


def _request_data() -> dict:
    """Extrai o corpo da requisicao (JSON ou form-data) e retorna como dicionario."""
    if request.is_json:
        return request.get_json(silent=True) or {}
    return request.form.to_dict(flat=True)


def _query_all(consulta: str, parametros=None) -> list:
    """Executa um SELECT e retorna todos os registros serializados como lista de dicionarios.

    Usa o conector mysql-connector se disponivel, caso contrario faz fallback para o CLI.
    """
    if mysql is None:
        return _query_all_cli(consulta, parametros)

    conexao = _connect_db()
    cursor = conexao.cursor(dictionary=True)
    try:
        cursor.execute(consulta, parametros or ())
        registros = cursor.fetchall()
        return [_serialize_row(r) for r in registros]
    finally:
        cursor.close()
        conexao.close()


def _query_one(consulta: str, parametros=None):
    """Executa um SELECT e retorna o primeiro registro serializado, ou None se vazio.

    Usa o conector mysql-connector se disponivel, caso contrario faz fallback para o CLI.
    """
    if mysql is None:
        registros = _query_all_cli(consulta, parametros)
        return registros[0] if registros else None

    conexao = _connect_db()
    cursor = conexao.cursor(dictionary=True)
    try:
        cursor.execute(consulta, parametros or ())
        registro = cursor.fetchone()
        return _serialize_row(registro) if registro else None
    finally:
        cursor.close()
        conexao.close()


def _execute_write(consulta: str, parametros=None):
    """Executa uma operacao de escrita (INSERT/UPDATE/DELETE) dentro de uma unica transacao.

    Faz rollback automatico em caso de excecao e sempre fecha a conexao.
    Retorna (lastrowid, rowcount).
    """
    conexao = _connect_db()
    cursor = conexao.cursor()
    try:
        cursor.execute(consulta, parametros or ())
        conexao.commit()
        return cursor.lastrowid, cursor.rowcount
    except Exception:
        conexao.rollback()
        raise
    finally:
        cursor.close()
        conexao.close()


@app.route("/")
def home():
    """Entrega a pagina de login como rota raiz padrao da aplicacao."""
    return _serve_html_with_dev_reload(os.path.join(DIRETORIO_FRONTEND, "login.html"))


@app.route("/<path:caminho>")
def static_files(caminho):
    """Entrega arquivos estaticos do frontend, injetando auto-reload em paginas HTML."""
    # Tenta primeiro na pasta frontend/
    alvo_frontend = os.path.join(DIRETORIO_FRONTEND, caminho)
    if os.path.isfile(alvo_frontend):
        if caminho.lower().endswith(".html"):
            return _serve_html_with_dev_reload(alvo_frontend)
        return send_from_directory(DIRETORIO_FRONTEND, caminho)

    # Fallback: tenta na raiz do projeto (ex: arquivos de impressao, modelos)
    alvo_projeto = os.path.join(RAIZ_PROJETO, caminho)
    if caminho.lower().endswith(".html") and os.path.isfile(alvo_projeto):
        return _serve_html_with_dev_reload(alvo_projeto)
    return send_from_directory(RAIZ_PROJETO, caminho)


@app.after_request
def compress_json_response(response):
    """Aplica compressao gzip em respostas JSON elegiveis para reduzir o tamanho do payload."""
    try:
        if response is None or response.direct_passthrough:
            return response
        if response.status_code < 200 or response.status_code >= 300:
            return response
        if response.headers.get("Content-Encoding"):
            return response

        tipo_conteudo = (response.headers.get("Content-Type") or "").lower()
        if "application/json" not in tipo_conteudo:
            return response

        aceita_codificacao = (request.headers.get("Accept-Encoding") or "").lower()
        if "gzip" not in aceita_codificacao:
            return response
        # /api/fichas pode retornar payload muito grande — pula compressao para evitar picos de memoria
        if request.path == "/api/fichas":
            return response

        dados_brutos = response.get_data()
        if not dados_brutos or len(dados_brutos) < 1024:
            return response
        if len(dados_brutos) > LIMITE_BYTES_GZIP:
            return response

        dados_comprimidos = gzip.compress(dados_brutos, compresslevel=5)
        if len(dados_comprimidos) >= len(dados_brutos):
            return response

        response.set_data(dados_comprimidos)
        response.headers["Content-Encoding"] = "gzip"
        response.headers["Content-Length"] = str(len(dados_comprimidos))
        response.headers["Vary"] = "Accept-Encoding"
        return response
    except Exception:
        return response


@app.route("/api/fichas", methods=["POST"])
def salvar_ficha():
    """Guarda somente leitura — endpoint de criacao desabilitado."""
    try:
        return jsonify(
            {
                "success": False,
                "message": "API em modo somente leitura. Criacao de ficha desabilitada.",
            }
        ), 405
    except (MySQLError, RuntimeError, ValueError) as error:
        return jsonify({"success": False, "message": str(error)}), 400


@app.route("/api/fichas", methods=["GET"])
def buscar_fichas():
    """Lista fichas com filtros opcionais e cache de curta duracao para reducao de carga no banco."""
    try:
        busca = request.args.get("search", "").strip().lower()
        tipo = request.args.get("tipo", "").strip()
        status = request.args.get("status", "").strip()
        data_inicio = request.args.get("data_inicio", "").strip()
        data_fim = request.args.get("data_fim", "").strip()
        # Ativa cache apenas quando ha filtros — listagem completa nunca e cacheada
        usar_cache = any([busca, tipo, status, data_inicio, data_fim])
        if usar_cache:
            chave_cache = ("fichas", busca, tipo, status, data_inicio, data_fim)
            dados_cache = _cache_get(chave_cache)
            if dados_cache is not None:
                resposta = jsonify(dados_cache)
                resposta.headers["X-Data-Cache"] = "HIT"
                return resposta, 200

        consulta = """
            SELECT
                numero_ficha,
                REPLACE(REPLACE(REPLACE(COALESCE(job, ''), CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' ') AS job,
                REPLACE(REPLACE(REPLACE(COALESCE(job_envolvidas, ''), CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' ') AS job_envolvidas,
                REPLACE(REPLACE(REPLACE(COALESCE(elemento_pep, ''), CHAR(13), '||'), CHAR(10), '||'), CHAR(9), ' ') AS elemento_pep,
                REPLACE(REPLACE(REPLACE(COALESCE(tipo_lamina_processo, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS tipo_lamina_processo,
                REPLACE(REPLACE(REPLACE(COALESCE(desenhos, ''), CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' ') AS desenhos,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(ficha_processo, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS ficha_processo,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_desenho_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_desenho_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(status_roteiro, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_roteiro,
                REPLACE(REPLACE(REPLACE(COALESCE(status_frnr, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_frnr,
                REPLACE(REPLACE(REPLACE(COALESCE(status_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(status_entrega_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_entrega_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(descricao_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS descricao_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(pt, ''), CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' ') AS pt,
                data_entrega_evento, data_emissao_requisicao,
                REPLACE(REPLACE(REPLACE(COALESCE(operador_responsavel, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS operador_responsavel,
                REPLACE(REPLACE(REPLACE(COALESCE(ferramenteiro_responsavel, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS ferramenteiro_responsavel,
                qtde_afiacao,
                data_termino_producao, data_termino_avaliacao_kit,
                REPLACE(REPLACE(REPLACE(COALESCE(observacao_ultima_manutencao_ferramenta, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS observacao_ultima_manutencao_ferramenta,
                REPLACE(REPLACE(REPLACE(COALESCE(observacao_rnc, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS observacao_rnc,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_plan_interno, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_plan_interno,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_po, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_po,
                vida_util_ferramenta,
                REPLACE(REPLACE(REPLACE(COALESCE(prensa, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS prensa,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_rnc, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_rnc,
                total_golpes_final, total_golpes_sem_afiar,
                `1_golpes_afiacao`, `2_golpes_afiacao`, `3_golpes_afiacao`, `4_golpes_afiacao`, `5_golpes_afiacao`,
                qtde_total_afiacao, qtde_total_golpes_ferramenta, qtde_uso_ferramenta,
                data_envio_ferramentaria, horario_inicio_ferramenteiro, horario_termino_ferramenteiro
            FROM banco_dados
            WHERE 1=1
        """
        parametros = []

        if busca:
            consulta += " AND (LOWER(COALESCE(job, '')) LIKE %s OR LOWER(COALESCE(job_envolvidas, '')) LIKE %s OR LOWER(COALESCE(numero_ferramental, '')) LIKE %s)"
            valor_like = f"%{busca}%"
            parametros.extend([valor_like, valor_like, valor_like])
        if tipo:
            consulta += " AND tipo_lamina_processo = %s"
            parametros.append(tipo)
        if status:
            consulta += " AND status_roteiro = %s"
            parametros.append(status)
        if data_inicio:
            consulta += " AND data_termino_producao >= %s"
            parametros.append(data_inicio)
        if data_fim:
            consulta += " AND data_termino_producao <= %s"
            parametros.append(data_fim)

        consulta += " ORDER BY numero_ficha DESC"
        registros = _query_all(consulta, parametros)
        dados = [_map_banco_dados_to_api(r) for r in registros]
        if usar_cache:
            _cache_set(chave_cache, dados, ttl_seconds=20)
        resposta = jsonify(dados)
        if usar_cache:
            resposta.headers["X-Data-Cache"] = "MISS"
        return resposta, 200
    except (MySQLError, RuntimeError) as error:
        return jsonify({"error": str(error)}), 400


@app.route("/api/estatisticas", methods=["GET"])
def obter_estatisticas():
    """Retorna contadores do dashboard agregados a partir da tabela banco_dados."""
    try:
        consulta = """
            SELECT
                COUNT(*) AS total,
                SUM(CASE WHEN LOWER(COALESCE(status_roteiro, '')) LIKE '%aprovado%' OR LOWER(COALESCE(status_roteiro, '')) LIKE '%conclu%' THEN 1 ELSE 0 END) AS concluidas,
                SUM(CASE WHEN LOWER(COALESCE(status_roteiro, '')) LIKE '%roteiro%' OR status_roteiro IS NULL OR LOWER(status_roteiro) = 'na' THEN 1 ELSE 0 END) AS pendentes,
                SUM(CASE WHEN NOT (LOWER(COALESCE(status_roteiro, '')) LIKE '%aprovado%' OR LOWER(COALESCE(status_roteiro, '')) LIKE '%conclu%' OR LOWER(COALESCE(status_roteiro, '')) LIKE '%roteiro%' OR status_roteiro IS NULL OR LOWER(status_roteiro) = 'na') THEN 1 ELSE 0 END) AS em_andamento,
                COALESCE(SUM(vida_util_ferramenta), 0) AS tempo_total
            FROM banco_dados
        """
        registro = _query_one(consulta) or {}
        estatisticas = {
            "total": int(registro.get("total", 0) or 0),
            "pendentes": int(registro.get("pendentes", 0) or 0),
            "em_andamento": int(registro.get("em_andamento", 0) or 0),
            "concluidas": int(registro.get("concluidas", 0) or 0),
            "tempo_total": float(registro.get("tempo_total", 0) or 0),
        }
        return jsonify(estatisticas), 200
    except (MySQLError, RuntimeError) as error:
        return jsonify({"error": str(error)}), 400


@app.route("/api/exportar-excel", methods=["GET"])
def exportar_excel():
    """Exporta o conjunto atual de fichas para um arquivo .xlsx para download."""
    try:
        fichas = _query_all(
            """
            SELECT
                numero_ficha, job, job_envolvidas, tipo_lamina_processo, numero_ferramental,
                status_roteiro, descricao_ferramental, pt, data_emissao_requisicao,
                operador_responsavel, ferramenteiro_responsavel, qtde_afiacao,
                data_termino_producao, observacao_ultima_manutencao_ferramenta
            FROM banco_dados
            ORDER BY numero_ficha DESC
            """
        )

        pasta_trabalho = openpyxl.Workbook()
        planilha = pasta_trabalho.active
        planilha.title = "Fichas de Manutencao"

        preenchimento_header = PatternFill(start_color="667eea", end_color="667eea", fill_type="solid")
        fonte_header = Font(color="FFFFFF", bold=True)
        borda = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        cabecalhos = [
            "Numero Ficha",
            "Job",
            "Jobs Envolvidas",
            "Tipo Lamina Processo",
            "Numero Ferramental",
            "Status Roteiro",
            "Descricao Ferramental",
            "PT",
            "Data Emissao Requisicao",
            "Operador Responsavel",
            "Ferramenteiro Responsavel",
            "Qtde Afiacao",
            "Data Termino Producao",
            "Observacao Ultima Manutencao",
        ]

        for col, cabecalho in enumerate(cabecalhos, 1):
            celula = planilha.cell(row=1, column=col, value=cabecalho)
            celula.fill = preenchimento_header
            celula.font = fonte_header
            celula.alignment = Alignment(horizontal="center", vertical="center")
            celula.border = borda

        for indice_linha, ficha in enumerate(fichas, 2):
            planilha.cell(row=indice_linha, column=1, value=ficha.get("numero_ficha"))
            planilha.cell(row=indice_linha, column=2, value=ficha.get("job"))
            planilha.cell(row=indice_linha, column=3, value=ficha.get("job_envolvidas"))
            planilha.cell(row=indice_linha, column=4, value=ficha.get("tipo_lamina_processo"))
            planilha.cell(row=indice_linha, column=5, value=ficha.get("numero_ferramental"))
            planilha.cell(row=indice_linha, column=6, value=ficha.get("status_roteiro"))
            planilha.cell(row=indice_linha, column=7, value=ficha.get("descricao_ferramental"))
            planilha.cell(row=indice_linha, column=8, value=ficha.get("pt"))
            planilha.cell(row=indice_linha, column=9, value=ficha.get("data_emissao_requisicao"))
            planilha.cell(row=indice_linha, column=10, value=ficha.get("operador_responsavel"))
            planilha.cell(row=indice_linha, column=11, value=ficha.get("ferramenteiro_responsavel"))
            planilha.cell(row=indice_linha, column=12, value=ficha.get("qtde_afiacao"))
            planilha.cell(row=indice_linha, column=13, value=ficha.get("data_termino_producao"))
            planilha.cell(row=indice_linha, column=14, value=ficha.get("observacao_ultima_manutencao_ferramenta"))

        for col in range(1, len(cabecalhos) + 1):
            planilha.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

        buffer_saida = BytesIO()
        pasta_trabalho.save(buffer_saida)
        buffer_saida.seek(0)

        nome_arquivo = f"Fichas_Manutencao_{datetime.now().strftime('%d_%m_%Y')}.xlsx"
        return send_file(
            buffer_saida,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=nome_arquivo,
        )
    except (MySQLError, RuntimeError) as error:
        return jsonify({"error": str(error)}), 400


@app.route("/api/ficha/<int:ficha_id>", methods=["GET"])
def obter_ficha(ficha_id):
    """Retorna uma unica ficha pelo id com suporte a cache de curta duracao."""
    try:
        chave_cache = ("ficha", int(ficha_id))
        dados_cache = _cache_get(chave_cache)
        if dados_cache is not None:
            resposta = jsonify(dados_cache)
            resposta.headers["X-Data-Cache"] = "HIT"
            return resposta, 200

        ficha = _query_one(
            """
            SELECT
                numero_ficha,
                REPLACE(REPLACE(REPLACE(COALESCE(job, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS job,
                REPLACE(REPLACE(REPLACE(COALESCE(job_envolvidas, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS job_envolvidas,
                REPLACE(REPLACE(REPLACE(COALESCE(tipo_lamina_processo, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS tipo_lamina_processo,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(ficha_processo, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS ficha_processo,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_desenho_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_desenho_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(status_roteiro, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_roteiro,
                REPLACE(REPLACE(REPLACE(COALESCE(status_frnr, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_frnr,
                REPLACE(REPLACE(REPLACE(COALESCE(status_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(status_entrega_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS status_entrega_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(descricao_ferramental, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS descricao_ferramental,
                REPLACE(REPLACE(REPLACE(COALESCE(pt, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS pt,
                data_entrega_evento, data_emissao_requisicao,
                REPLACE(REPLACE(REPLACE(COALESCE(operador_responsavel, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS operador_responsavel,
                REPLACE(REPLACE(REPLACE(COALESCE(ferramenteiro_responsavel, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS ferramenteiro_responsavel,
                qtde_afiacao,
                data_termino_producao, data_termino_avaliacao_kit,
                REPLACE(REPLACE(REPLACE(COALESCE(observacao_ultima_manutencao_ferramenta, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS observacao_ultima_manutencao_ferramenta,
                REPLACE(REPLACE(REPLACE(COALESCE(observacao_rnc, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS observacao_rnc,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_plan_interno, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_plan_interno,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_po, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_po,
                vida_util_ferramenta,
                REPLACE(REPLACE(REPLACE(COALESCE(prensa, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS prensa,
                REPLACE(REPLACE(REPLACE(COALESCE(numero_rnc, ''), '\r', ' '), '\n', ' '), '\t', ' ') AS numero_rnc,
                total_golpes_final, total_golpes_sem_afiar,
                `1_golpes_afiacao`, `2_golpes_afiacao`, `3_golpes_afiacao`, `4_golpes_afiacao`, `5_golpes_afiacao`,
                qtde_total_afiacao, qtde_total_golpes_ferramenta, qtde_uso_ferramenta,
                data_envio_ferramentaria, horario_inicio_ferramenteiro, horario_termino_ferramenteiro
            FROM banco_dados
            WHERE numero_ficha = %s
            """,
            (ficha_id,),
        )
        if not ficha:
            return jsonify({"error": "Ficha nao encontrada"}), 404
        dados = _map_banco_dados_to_api(ficha)
        _cache_set(chave_cache, dados, ttl_seconds=30)
        resposta = jsonify(dados)
        resposta.headers["X-Data-Cache"] = "MISS"
        return resposta, 200
    except (MySQLError, RuntimeError) as error:
        return jsonify({"error": str(error)}), 400


@app.route("/api/ficha/<int:ficha_id>", methods=["PUT"])
def atualizar_ficha(ficha_id):
    """Guarda somente leitura — endpoint de atualizacao desabilitado."""
    try:
        return jsonify(
            {
                "success": False,
                "message": "API em modo somente leitura. Atualizacao desabilitada.",
            }
        ), 405
    except (MySQLError, RuntimeError, ValueError) as error:
        return jsonify({"success": False, "message": str(error)}), 400


@app.route("/api/ficha/<int:ficha_id>", methods=["DELETE"])
def deletar_ficha(ficha_id):
    """Guarda somente leitura — endpoint de exclusao desabilitado."""
    try:
        return jsonify(
            {
                "success": False,
                "message": "API em modo somente leitura. Exclusao desabilitada.",
            }
        ), 405
    except (MySQLError, RuntimeError) as error:
        return jsonify({"success": False, "message": str(error)}), 400


@app.route("/health", methods=["GET"])
def health_check():
    """Verifica a conectividade com o banco de dados para monitoramento de saude do servico."""
    try:
        registro = _query_one("SELECT 1 AS ok")
        valor_ok = None if not registro else registro.get("ok")
        if str(valor_ok) == "1":
            return jsonify({"status": "ok", "storage": "mysql"}), 200
        return jsonify({"status": "error", "storage": "mysql"}), 500
    except Exception as error:
        return jsonify({"status": "error", "storage": "mysql", "message": str(error)}), 500


@app.route("/__dev_timestamp", methods=["GET"])
def dev_timestamp():
    """Retorna o timestamp do arquivo estatico mais recente; usado pelo auto-reload local."""
    try:
        diretorios_raiz = [DIRETORIO_FRONTEND, RAIZ_PROJETO]
        extensoes = {".html", ".css", ".js"}
        mais_recente = 0.0

        for diretorio in diretorios_raiz:
            if not os.path.isdir(diretorio):
                continue
            for pasta, _, arquivos in os.walk(diretorio):
                for nome in arquivos:
                    _, extensao = os.path.splitext(nome.lower())
                    if extensao not in extensoes:
                        continue
                    caminho = os.path.join(pasta, nome)
                    try:
                        mais_recente = max(mais_recente, os.path.getmtime(caminho))
                    except OSError:
                        pass

        return jsonify({"ts": int(mais_recente * 1000)}), 200
    except Exception as error:
        return jsonify({"error": str(error)}), 500


if __name__ == "__main__":
    # Vincula em todas as interfaces para o site ser acessivel da LAN e de tuneis
    host_servidor = os.getenv("FLASK_HOST", "0.0.0.0")
    porta = int(os.getenv("PORT", "5000"))
    modo_debug = os.getenv("FLASK_DEBUG", "0").strip().lower() in ("1", "true", "yes", "on")
    app.run(debug=modo_debug, port=porta, host=host_servidor, use_reloader=False)
