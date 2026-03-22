/*
 * Arquivo: frontend/script.js
 * Descricao: handlers globais de UI — submit de formularios, filtros e stats do dashboard.
 */

// URL base da API
const API_URL = 'http://localhost:5000/api';

// Gerenciamento de tabs (se existirem na tela)
document.querySelectorAll('.tab-btn').forEach(button => {
    button.addEventListener('click', () => {
        const nomeAba = button.getAttribute('data-tab');
/*
 * Notas de manutencao:
 * - Objetivo: preservar interacoes de UI e chamadas de API.
 * - Cuidado: validar elementos por id/classe antes de alterar seletores.
 * - Ao alterar: testar filtros, listagens e mensagens de erro.
 */

        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        button.classList.add('active');
        document.getElementById(nomeAba).classList.add('active');
        if (nomeAba === 'relatorio') carregarEstatisticas();
    });
});

// Auto-preencher datas do formulario
document.addEventListener('DOMContentLoaded', () => {
    preencherDatasFormulario();
    agendarAtualizacaoDiaria();
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) preencherDatasFormulario();
    });
});

// Submit do formulário (cria ficha)
document.getElementById('formNovaFicha').addEventListener('submit', async (e) => {
    e.preventDefault();
    const dadosFormulario = new FormData(e.target);

    try {
        const resposta = await fetch(`${API_URL}/fichas`, {
            method: 'POST',
            body: dadosFormulario
        });
        const resultado = await resposta.json();
        if (resultado.success) {
            mostrarMensagem('Ficha salva com sucesso!', 'success');
            e.target.reset();
            preencherDatasFormulario();
        } else {
            mostrarMensagem('Erro ao salvar: ' + resultado.message, 'error');
        }
    } catch (error) {
        mostrarMensagem('Erro ao comunicar com servidor', 'error');
    }
});

// Buscar fichas com filtros
async function buscarFichas() {
    const termoBusca = document.getElementById('searchInput')?.value || '';
    const status = document.getElementById('filtroStatus')?.value || '';
    const dataInicio = document.getElementById('filtroDataInicio')?.value || '';
    const dataFim = document.getElementById('filtroDataFim')?.value || '';

    const parametros = new URLSearchParams({search: termoBusca, status, data_inicio: dataInicio, data_fim: dataFim});

    try {
        const resposta = await fetch(`${API_URL}/fichas?${parametros}`);
        const fichas = await resposta.json();
        exibirFichas(fichas);
    } catch (error) {
        mostrarMensagem('Erro ao buscar fichas', 'error');
    }
}

// Exibir fichas na tela
function exibirFichas(fichas) {
    const containerResultados = document.getElementById('resultadosFichas');
    if (!containerResultados) return;

    if (fichas.length === 0) {
        containerResultados.innerHTML = '<p class="info-message">Nenhuma ficha encontrada</p>';
        return;
    }

    // Renderiza os cards de fichas com campos principais da manutencao
    containerResultados.innerHTML = fichas.map(ficha => `
        <div class="ficha-card">
            <div class="ficha-header">
                <div>
                    <h3>${ficha.descricao_ferramenta || 'Sem descrição'}</h3>
                    <small>Código: ${ficha.codigo_ferramenta}</small>
                </div>
                <span class="status-badge status-${ficha.status}">${formatarStatus(ficha.status)}</span>
            </div>
            <div class="ficha-info">
                <div class="info-item"><span class="info-label">Nº Ficha:</span><span class="info-value">${ficha.id}</span></div>
                <div class="info-item"><span class="info-label">Documento:</span><span class="info-value">${ficha.documento}</span></div>
                <div class="info-item"><span class="info-label">Área:</span><span class="info-value">${ficha.area}</span></div>
                <div class="info-item"><span class="info-label">Data Cadastro:</span><span class="info-value">${formatarData(ficha.data_cadastro)}</span></div>
                <div class="info-item"><span class="info-label">Responsável:</span><span class="info-value">${ficha.responsavel_abertura}</span></div>
                <div class="info-item"><span class="info-label">Ferramenteiro:</span><span class="info-value">${ficha.ferramenteiro}</span></div>
                <div class="info-item"><span class="info-label">Afiações:</span><span class="info-value">${ficha.qtde_afiacao || 0}</span></div>
                <div class="info-item"><span class="info-label">Data Término:</span><span class="info-value">${formatarData(ficha.data_termino)}</span></div>
            </div>
            <div style="margin-top: 15px;"><strong>Trabalhos:</strong> ${ficha.descricao_trabalhos}</div>
            ${ficha.observacoes ? `<div style="margin-top: 10px;"><strong>Observações:</strong> ${ficha.observacoes}</div>` : ''}
            <div style="margin-top: 15px;"><button onclick="deletarFicha(${ficha.id})" class="btn btn-danger">Deletar</button></div>
        </div>
    `).join('');
}

// Estatísticas do painel
async function carregarEstatisticas() {
    try {
        const resposta = await fetch(`${API_URL}/estatisticas`);
        const estatisticas = await resposta.json();
        document.getElementById('totalFichas').textContent = estatisticas.total || 0;
        document.getElementById('totalAndamento').textContent = estatisticas.em_andamento || 0;
        document.getElementById('totalPendentes').textContent = estatisticas.pendentes || 0;
        document.getElementById('totalConcluidas').textContent = estatisticas.concluidas || 0;
    } catch (error) {
        console.error('Erro ao carregar estatísticas:', error);
    }
}

// Exportação para Excel
async function gerarRelatorioExcel() {
    try {
        window.location.href = `${API_URL}/exportar-excel`;
        mostrarMensagem('Gerando Excel...', 'success');
    } catch (error) {
        mostrarMensagem('Erro ao gerar relatório', 'error');
    }
}

// Deletar ficha
async function deletarFicha(fichaId) {
    if (!confirm('Deletar esta ficha?')) return;
    try {
        const resposta = await fetch(`${API_URL}/ficha/${fichaId}`, {method: 'DELETE'});
        const resultado = await resposta.json();
        if (resultado.success) {
            mostrarMensagem('Ficha deletada!', 'success');
            buscarFichas();
        } else {
            mostrarMensagem('Erro: ' + resultado.message, 'error');
        }
    } catch (error) {
        mostrarMensagem('Erro ao comunicar', 'error');
    }
}

// Funções auxiliares
function mostrarMensagem(mensagem, tipo) {
    const caixaMensagem = document.getElementById('messageBox');
    if (!caixaMensagem) return;
    caixaMensagem.textContent = mensagem;
    caixaMensagem.className = `message-box message-${tipo}`;
    caixaMensagem.style.display = 'block';
    setTimeout(() => {caixaMensagem.style.display = 'none';}, 3000);
}

// Formatar status da ficha para texto legivel em pt-BR
function formatarStatus(status) {
    const mapaStatus = {'pendente': 'Pendente', 'em_andamento': 'Em Andamento', 'concluida': 'Concluída'};
    return mapaStatus[status] || status;
}

// Formatar data ISO para DD/MM/AAAA (exibicao em tabelas e cards)
function formatarData(textoData) {
    if (!textoData) return '';
    return new Date(textoData).toLocaleDateString('pt-BR');
}

function formatarDataFormulario(dataObj) {
    // Mantem o mesmo formato textual usado historicamente nas telas
    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    const dia = dataObj.getDate();
    const mes = meses[dataObj.getMonth()];
    const ano = dataObj.getFullYear();
    return `${dia}-${mes}-${ano}`;
}

// Preencher campos de data com a data atual ao carregar ou reativar a pagina
function preencherDatasFormulario() {

    const dataDocumento = document.querySelector('input[name="data"]');
    if (dataDocumento) dataDocumento.value = dataTexto;

    const dataCadastro = document.getElementById('data_cadastro');
    if (dataCadastro) dataCadastro.value = dataTexto;
}

// Reagendar preenchimento de datas na virada de meia-noite
function agendarAtualizacaoDiaria() {
    const agora = new Date();
    const proximaMeiaNoite = new Date(agora);
    proximaMeiaNoite.setHours(24, 0, 0, 0);
    const msAteMeiaNoite = proximaMeiaNoite.getTime() - agora.getTime();

    setTimeout(() => {
        preencherDatasFormulario();
        agendarAtualizacaoDiaria();
    }, msAteMeiaNoite);
}

// Listeners de filtros
document.getElementById('filtroStatus')?.addEventListener('change', buscarFichas);
document.getElementById('filtroDataInicio')?.addEventListener('change', buscarFichas);
document.getElementById('filtroDataFim')?.addEventListener('change', buscarFichas);
