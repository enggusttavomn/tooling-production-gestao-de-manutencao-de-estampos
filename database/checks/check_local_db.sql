-- Arquivo: database/check_local_db.sql
-- Descricao: script SQL do projeto Fichas de Manutencao.
-- Observacao: cabecalho de documentacao para padronizacao e organizacao.

SELECT NOW() AS agora, DATABASE() AS banco;
SELECT COUNT(*) AS total_registros FROM banco_dados;
SELECT job_envolvidas FROM banco_dados LIMIT 50;


-- Notas de manutencao:
-- - Objetivo: preservar integridade de dados e compatibilidade de schema.
-- - Cuidado: documentar impacto de alteracoes em colunas e constraints.
-- - Ao alterar: validar em base local antes de aplicar em ambiente compartilhado.

