-- Arquivo: insert_test_case.sql
-- Descricao: script SQL do projeto Fichas de Manutencao.
-- Observacao: cabecalho de documentacao para padronizacao e organizacao.

-- Inserir ficha de teste pendente (sem nenhum campo preenchido)
INSERT INTO banco_dados (
  numero_ficha,
  job,
  job_envolvidas,
  tipo_lamina_processo,
  status_roteiro,
  numero_ferramental,
-- Notas de manutencao:
-- - Objetivo: preservar integridade de dados e compatibilidade de schema.
-- - Cuidado: documentar impacto de alteracoes em colunas e constraints.
-- - Ao alterar: validar em base local antes de aplicar em ambiente compartilhado.

  descricao_ferramental,
  status_ferramental
) VALUES (
  99999,
  'JOB-TEST-001',
  'JOB-TEST-001',
  'LAMINA TESTE',
  'NA',
  'FERR-001',
  'Ferramenta de Teste',
  'A receber'
);

-- Verificar que foi inserido
SELECT numero_ficha, job, status_roteiro, total_golpes_final, qtde_afiacao FROM banco_dados WHERE numero_ficha = 99999;

