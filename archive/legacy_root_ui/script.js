const API_URL = 'http://localhost:5000/api';

// Gerenciamento de Tabs
document.querySelectorAll('.tab-btn').forEach(button => {
    button.addEventListener('click', () => {
        const tabName = button.getAttribute('data-tab');
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        button.classList.add('active');
        document.getElementById(tabName).classList.add('active');
        if (tabName === 'relatorio') carregarEstatisticas();
    });
});

// Auto-preencher data do cadastro
document.addEventListener('DOMContentLoaded', () => {
    preencherDatasFormulario();
    agendarAtualizacaoDiaria();
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) preencherDatasFormulario();
    });
});

// Submit do formulário
document.getElementById('formNovaFicha').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    
    try {
        const response = await fetch(`${API_URL}/fichas`, {
            method: 'POST',
            body: formData
        });
        const result = await response.json();
        if (result.success) {
            showMessage('Ficha salva com sucesso!', 'success');
            e.target.reset();
            preencherDatasFormulario();
        } else {
            showMessage('Erro ao salvar: ' + result.message, 'error');
        }
    } catch (error) {
        showMessage('Erro ao comunicar com servidor', 'error');
    }
});

// Buscar fichas
async function buscarFichas() {
    const searchTerm = document.getElementById('searchInput').value;
    const status = document.getElementById('filtroStatus').value;
    const dataInicio = document.getElementById('filtroDataInicio').value;
    const dataFim = document.getElementById('filtroDataFim').value;
    
    const params = new URLSearchParams({search: searchTerm, status, data_inicio: dataInicio, data_fim: dataFim});
    
    try {
        const response = await fetch(`${API_URL}/fichas?${params}`);
        const fichas = await response.json();
        exibirFichas(fichas);
    } catch (error) {
        showMessage('Erro ao buscar fichas', 'error');
    }
}

// Exibir fichas
function exibirFichas(fichas) {
    const container = document.getElementById('resultadosFichas');
    if (fichas.length === 0) {
        container.innerHTML = '<p class="info-message">Nenhuma ficha encontrada</p>';
        return;
    }
    
    container.innerHTML = fichas.map(ficha => `
        <div class="ficha-card">
            <div class="ficha-header">
                <div>
                    <h3>${ficha.descricao_ferramenta || 'Sem descrição'}</h3>
                    <small>Código: ${ficha.codigo_ferramenta}</small>
                </div>
                <span class="status-badge status-${ficha.status}">${formatStatus(ficha.status)}</span>
            </div>
            <div class="ficha-info">
                <div class="info-item"><span class="info-label">Nº Ficha:</span><span class="info-value">${ficha.id}</span></div>
                <div class="info-item"><span class="info-label">Documento:</span><span class="info-value">${ficha.documento}</span></div>
                <div class="info-item"><span class="info-label">Área:</span><span class="info-value">${ficha.area}</span></div>
                <div class="info-item"><span class="info-label">Data Cadastro:</span><span class="info-value">${formatDate(ficha.data_cadastro)}</span></div>
                <div class="info-item"><span class="info-label">Responsável:</span><span class="info-value">${ficha.responsavel_abertura}</span></div>
                <div class="info-item"><span class="info-label">Ferramenteiro:</span><span class="info-value">${ficha.ferramenteiro}</span></div>
                <div class="info-item"><span class="info-label">Afiações:</span><span class="info-value">${ficha.qtde_afiacao || 0}</span></div>
                <div class="info-item"><span class="info-label">Data Término:</span><span class="info-value">${formatDate(ficha.data_termino)}</span></div>
            </div>
            <div style="margin-top: 15px;"><strong>Trabalhos:</strong> ${ficha.descricao_trabalhos}</div>
            ${ficha.observacoes ? `<div style="margin-top: 10px;"><strong>Observações:</strong> ${ficha.observacoes}</div>` : ''}
            <div style="margin-top: 15px;"><button onclick="deletarFicha(${ficha.id})" class="btn btn-danger">Deletar</button></div>
        </div>
    `).join('');
}

// Estatísticas
async function carregarEstatisticas() {
    try {
        const response = await fetch(`${API_URL}/estatisticas`);
        const stats = await response.json();
        document.getElementById('totalFichas').textContent = stats.total || 0;
        document.getElementById('totalAndamento').textContent = stats.em_andamento || 0;
        document.getElementById('totalPendentes').textContent = stats.pendentes || 0;
        document.getElementById('totalConcluidas').textContent = stats.concluidas || 0;
    } catch (error) {
        console.error('Erro ao carregar estatísticas:', error);
    }
}

// Excel
async function gerarRelatorioExcel() {
    try {
        window.location.href = `${API_URL}/exportar-excel`;
        showMessage('Gerando Excel...', 'success');
    } catch (error) {
        showMessage('Erro ao gerar relatório', 'error');
    }
}

// Deletar
async function deletarFicha(fichaId) {
    if (!confirm('Deletar esta ficha?')) return;
    try {
        const response = await fetch(`${API_URL}/ficha/${fichaId}`, {method: 'DELETE'});
        const result = await response.json();
        if (result.success) {
            showMessage('Ficha deletada!', 'success');
            buscarFichas();
        } else {
            showMessage('Erro: ' + result.message, 'error');
        }
    } catch (error) {
        showMessage('Erro ao comunicar', 'error');
    }
}

// Auxiliares
function showMessage(message, type) {
    const messageBox = document.getElementById('messageBox');
    messageBox.textContent = message;
    messageBox.className = `message-box message-${type}`;
    messageBox.style.display = 'block';
    setTimeout(() => {messageBox.style.display = 'none';}, 3000);
}

function formatStatus(status) {
    const statusMap = {'pendente': 'Pendente', 'em_andamento': 'Em Andamento', 'concluida': 'Concluída'};
    return statusMap[status] || status;
}

function formatDate(dateString) {
    if (!dateString) return '';
    return new Date(dateString).toLocaleDateString('pt-BR');
}

function formatarDataFormulario(dateObj) {
    const meses = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const dia = dateObj.getDate();
    const mes = meses[dateObj.getMonth()];
    const ano = dateObj.getFullYear();
    return `${dia}-${mes}-${ano}`;
}

function preencherDatasFormulario() {
    const hoje = new Date();
    const dataTexto = formatarDataFormulario(hoje);

    const dataDocumento = document.querySelector('input[name="data"]');
    if (dataDocumento) dataDocumento.value = dataTexto;

    const dataCadastro = document.getElementById('data_cadastro');
    if (dataCadastro) dataCadastro.value = dataTexto;
}

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

// Listeners
document.getElementById('filtroStatus')?.addEventListener('change', buscarFichas);
document.getElementById('filtroDataInicio')?.addEventListener('change', buscarFichas);
document.getElementById('filtroDataFim')?.addEventListener('change', buscarFichas);
