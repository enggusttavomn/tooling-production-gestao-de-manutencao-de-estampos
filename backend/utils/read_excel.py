"""
Arquivo: backend/utils/read_excel.py
Descricao: utilitario de diagnostico para inspecionar celulas preenchidas no Excel.

Uso tipico:
- Validar rapidamente se a planilha base foi carregada corretamente.
- Conferir coordenadas e valores sem abrir o arquivo manualmente no Excel.
"""

from pathlib import Path

import openpyxl
# Notas de manutencao:
# - Objetivo: manter comportamento funcional sem quebrar compatibilidade.
# - Cuidado: validar efeitos em fluxo web, banco e exportacoes.
# - Ao alterar: preferir mudancas pequenas e validacao local.



DEFAULT_WORKBOOK = "Ficha de Manutenção Padrão Rev. - 05_Julho_2024.xlsx"


def imprimir_celulas_preenchidas(caminho_arquivo: Path, max_linhas: int = 34, max_colunas: int = 5) -> None:
    """Imprime coordenadas/valores de celulas nao vazias no intervalo informado."""
    pasta_trabalho = openpyxl.load_workbook(caminho_arquivo)
    planilha = pasta_trabalho.active

    for num_linha in range(1, max_linhas + 1):
        for num_coluna in range(1, max_colunas + 1):
            celula = planilha.cell(num_linha, num_coluna)
            if celula.value not in (None, ""):
                print(f"{celula.coordinate}: {celula.value}")


if __name__ == "__main__":
    # O arquivo padrao esta centralizado em backend/data apos reorganizacao.
    diretorio_backend = Path(__file__).resolve().parents[1]
    caminho_arquivo = diretorio_backend / "data" / DEFAULT_WORKBOOK
    imprimir_celulas_preenchidas(caminho_arquivo)
