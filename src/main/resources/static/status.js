/**
 * @file status.js
 * @description Gerencia a lógica do modal de status, gráfico e legendas editáveis.
 */

const KEY_LEGENDAS = 'legendaNiveis';

/**
 * Carrega as legendas salvas do localStorage.
 * @returns {Promise<object>} Um objeto com as legendas.
 */
async function carregarLegendasSalvas() {
    try {
        const legendasSalvas = localStorage.getItem(KEY_LEGENDAS);
        if (legendasSalvas) {
            return JSON.parse(legendasSalvas);
        }
    } catch (error) {
        console.error("Erro ao carregar legendas:", error);
    }
    // Retorna um objeto padrão se não houver nada salvo
    const legendasPadrao = {};
    for (let i = 1; i <= 10; i++) {
        legendasPadrao[i] = `Nível ${i}`;
    }
    return legendasPadrao;
}

/**
 * Salva uma legenda específica no localStorage.
 * @param {number} nivel - O nível da legenda (1-10).
 * @param {string} texto - O novo texto da legenda.
 */
async function salvarLegenda(nivel, texto) {
    try {
        const legendasAtuais = await carregarLegendasSalvas();
        legendasAtuais[nivel] = texto;
        localStorage.setItem(KEY_LEGENDAS, JSON.stringify(legendasAtuais));
    } catch (error) {
        console.error("Erro ao salvar legenda:", error);
    }
}


/**
 * Calcula os percentuais e gera o gráfico de barras com legendas editáveis.
 */
async function abrirModalGrafico() {
    const chartContainer = document.getElementById('chartContainer');
    const legendContainer = document.getElementById('chart-legend');
    const modal = document.getElementById('modalGrafico');
    if (!chartContainer || !modal || !legendContainer) return;

    const legendas = await carregarLegendasSalvas();
    const counts = {};
    let totalLinhas = 0;

    for (let i = 1; i <= 10; i++) {
        const countSpan = document.getElementById(`count-nivel-${i}`);
        const count = countSpan ? parseInt(countSpan.textContent, 10) : 0;
        counts[i] = count;
        totalLinhas += count;
    }

    chartContainer.innerHTML = '';
    legendContainer.innerHTML = '';

    if (totalLinhas === 0) {
        chartContainer.innerHTML = '<p>Não há linhas na tabela para gerar o gráfico.</p>';
        modal.style.display = 'flex';
        return;
    }

    for (let i = 1; i <= 10; i++) {
        const count = counts[i];
        const percent = totalLinhas > 0 ? ((count / totalLinhas) * 100) : 0;
        const corNivel = `var(--nivel-${i}-color)`;

        // [CORRIGIDO] Cria a legenda à esquerda
        const legendItem = document.createElement('div');
        legendItem.className = 'legend-item';
        legendItem.innerHTML = `
            <div class="legend-color-dot" style="background-color: ${corNivel};"></div>
            <input type="text" class="legend-input" id="legend-input-${i}" value="${legendas[i] || `Nível ${i}`}" style="color: ${corNivel};">
        `;
        legendContainer.appendChild(legendItem);

        const legendInput = document.getElementById(`legend-input-${i}`);
        legendInput.addEventListener('change', (e) => {
            salvarLegenda(i, e.target.value);
        });

        // [CORRIGIDO] Cria a barra do gráfico à direita
        const barWrapper = document.createElement('div');
        barWrapper.className = 'chart-bar-wrapper';
        barWrapper.innerHTML = `
            <div class="bar-percentage">${percent.toFixed(1)}%</div>
            <div class="chart-bar" style="height: ${percent.toFixed(1)}%; background-color: ${corNivel};" title="${legendas[i]}: ${count} linha(s) - ${percent.toFixed(1)}%"></div>
            <div class="bar-label">Nível ${i}</div>
        `;
        chartContainer.appendChild(barWrapper);
    }

    modal.style.display = 'flex';
}

/**
 * Adiciona os listeners de eventos para o modal do gráfico assim que a página carregar.
 */
document.addEventListener("DOMContentLoaded", () => {
    const graficoNiveisBtn = document.getElementById('graficoNiveisBtn');
    const modalGrafico = document.getElementById('modalGrafico');
    const fecharModalGraficoBtn = document.getElementById('fecharModalGraficoBtn');

    if (graficoNiveisBtn) {
        graficoNiveisBtn.addEventListener('click', abrirModalGrafico);
    }
    if (fecharModalGraficoBtn) {
        fecharModalGraficoBtn.addEventListener('click', () => modalGrafico.style.display = 'none');
    }
    if (modalGrafico) {
        modalGrafico.addEventListener('click', (e) => {
            if (e.target === modalGrafico) modalGrafico.style.display = 'none';
        });
    }
});