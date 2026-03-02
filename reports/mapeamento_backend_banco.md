# Mapeamento inicial backend x banco_dados

## Rotas ja ligadas ao banco real
- GET /api/fichas -> SELECT em app_ferramental.banco_dados
- GET /api/ficha/<id> -> numero_ficha
- GET /api/estatisticas -> agregacoes por status_roteiro
- GET /api/exportar-excel -> exporta dados de banco_dados

## Campos mapeados para resposta da API
| Campo API | Coluna banco_dados |
|---|---|
| id / numero_ficha | numero_ficha |
| equipamento / tipo_manutencao | tipo_lamina_processo |
| tag / codigo_ferramenta | numero_ferramental |
| status | status_roteiro (normalizado) |
| descricao / descricao_ferramenta | descricao_ferramental |
| area | pt |
| data_cadastro / data_criacao | data_emissao_requisicao |
| responsavel / responsavel_abertura | operador_responsavel |
| ferramenteiro | ferramenteiro_responsavel |
| qtde_afiacao | qtde_afiacao |
| data_termino / data_atualizacao | data_termino_producao |
| observacoes | observacao_ultima_manutencao_ferramenta ou observacao_rnc |
| documento | numero_plan_interno ou numero_po |
| job | job |
| job_envolvidas | job_envolvidas |

## Pendente (proxima etapa)
- POST /api/fichas: definir mapeamento oficial dos campos do formulario para INSERT em banco_dados.
- PUT /api/ficha/<id>: definir quais colunas podem ser alteradas por perfil.
- Validar trigger trg_numero_ficha no ambiente local/remoto para insercao.
