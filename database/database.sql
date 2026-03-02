-- Criar banco de dados
CREATE DATABASE IF NOT EXISTS fichas_manutencao CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

-- Selecionar banco
USE fichas_manutencao;

-- Tabela de fichas de manutencao
CREATE TABLE IF NOT EXISTS fichas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    equipamento VARCHAR(255) NOT NULL,
    tag VARCHAR(100) NOT NULL,
    tipo_manutencao ENUM('preventiva', 'corretiva', 'preditiva') NOT NULL,
    data_execucao DATE NOT NULL,
    responsavel VARCHAR(255) NOT NULL,
    area VARCHAR(255) NOT NULL,
    descricao TEXT NOT NULL,
    observacoes TEXT,
    tempo_execucao DECIMAL(5,2) NOT NULL,
    status ENUM('pendente', 'em_andamento', 'concluida') DEFAULT 'pendente',
    data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_equipamento (equipamento),
    INDEX idx_tag (tag),
    INDEX idx_tipo (tipo_manutencao),
    INDEX idx_status (status),
    INDEX idx_data_execucao (data_execucao)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- Tabela de historico de alteracoes
CREATE TABLE IF NOT EXISTS historico_alteracoes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    ficha_id INT NOT NULL,
    campo_alterado VARCHAR(100) NOT NULL,
    valor_anterior TEXT,
    valor_novo TEXT,
    usuario VARCHAR(255),
    data_alteracao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ficha_id) REFERENCES fichas(id) ON DELETE CASCADE,
    INDEX idx_ficha (ficha_id),
    INDEX idx_data (data_alteracao)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- Inserir alguns dados de exemplo
INSERT INTO fichas (equipamento, tag, tipo_manutencao, data_execucao, responsavel, area, descricao, tempo_execucao, status) VALUES
('Bomba Centrifuga', 'BC-001', 'preventiva', '2026-02-10', 'Joao Silva', 'Utilidades', 'Lubrificacao e inspecao de rolamentos', 2.5, 'pendente'),
('Compressor de Ar', 'CA-015', 'corretiva', '2026-02-06', 'Maria Santos', 'Producao', 'Substituicao de valvula de seguranca', 4.0, 'em_andamento'),
('Trocador de Calor', 'TC-023', 'preventiva', '2026-02-15', 'Pedro Costa', 'Processo', 'Limpeza quimica e teste de pressao', 6.5, 'pendente'),
('Motor Eletrico', 'ME-045', 'preditiva', '2026-02-08', 'Ana Oliveira', 'Manutencao', 'Analise de vibracao e termografia', 3.0, 'concluida');
