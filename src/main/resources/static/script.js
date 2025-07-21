/**
 * @file script.js
 * @description Script principal para a aplicação de Lista Técnica IFS.
 * Gerencia a interação com a tabela, importação/exportação de Excel,
 * autopreenchimento de colunas, e funcionalidades de UI.
 */

/**
 * Referência ao corpo (tbody) da tabela HTML onde os dados serão exibidos.
 * @type {HTMLTableSectionElement}
 */
let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];

/**
 * Cache para armazenar linhas copiadas, permitindo a operação de colar.
 * @type {Array<Object>}
 */
let cacheCopiado = [];

/**
 * Estado de visibilidade da coluna "SEQ" (Sequência).
 * @type {boolean}
 */
let seqAtivo = true;

/**
 * Estado de visibilidade da coluna "NÍVEL".
 * @type {boolean}
 */
let nivelColVisivel = true;

/**
 * Cor hexadecimal atualmente selecionada para demarcação de células/linhas.
 * @type {string}
 */
let corSelecionada = "";

/**
 * Modo de demarcação: true para demarcar a linha inteira, false para a célula.
 * @type {boolean}
 */
let demarcarLinha = false;

/**
 * Modo de remoção de demarcação: true para remover cores ao clicar, false para aplicar.
 * @type {boolean}
 */
let removerDemarcacao = false;

/**
 * Estado para ignorar ou aplicar o destaque de linhas duplicadas.
 * @type {boolean}
 */
let ignorarDuplicatas = false;

/**
 * Estado de ativação do efeito de "régua" (highlight da linha ao passar o mouse).
 * @type {boolean}
 */
let hoverEffectAtivo = true;

/**
 * Definição das cores em formato hexadecimal para cada nível de indentação.
 * @type {Array<string>}
 */
const nivelColors = [
    "#4664cf", "#CD5C5C", "#B3E6B3", "#FFD700", "#8A2BE2",
    "#FF8C00", "#00CED1", "#FF69B4", "#9ACD32", "#DA70D6"
];

/**
 * Opções para a coluna "TIPO ESTRUTURA".
 * @type {Array<string>}
 */
const tiposEstrutura = ["Manufatura", "Comprado", ""];

/**
 * Opções para a coluna "FATOR_SUCATA".
 * @type {Array<string>}
 */
const fatorSucata = ["0", "15", ""];

/**
 * Opções para a coluna "ALTERNATIVA".
 * @type {Array<string>}
 */
const alternativas = ["*", ""];

/**
 * Opções para a coluna "SITE".
 * @type {Array<string>}
 */
const siteValores = ["1", ""];

/**
 * Níveis de 1 a 10 para a coluna "NÍVEL".
 * @type {Array<string>}
 */
const niveis = Array.from({ length: 10 }, (_, i) => (i + 1).toString());

// --- Variáveis de estado para a navegação de duplicatas ---
let foundDuplicates = []; // Armazena as TRs que são duplicatas
let currentDuplicateIndex = -1; // Índice da duplicata atualmente focada

/**
 * Reseta o estado da busca de duplicatas e remove destaques de foco.
 */
function resetDuplicateSearchState() {
    foundDuplicates = [];
    currentDuplicateIndex = -1;
    // Remove as classes de opacidade e destaque de todas as linhas e do tbody
    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
    tbodyElement.classList.remove("table-faded");
    tabela.querySelectorAll('tr.highlight-focused-item').forEach(row => {
        row.classList.remove('highlight-focused-item');
        row.classList.remove('temp-highlight-found');
    });
    // Força uma nova verificação de duplicatas para re-aplicar o roxo claro em todos os itens que ainda são duplicatas
    verificarDuplicatas();
}

/**
 * Procura por uma linha existente na tabela com o mesmo CODIGO_MATERIAL e ITEM_COMPONENTE.
 * Ignora a própria linha se for passada para edição.
 * @param {string} codigoMaterial - O CODIGO_MATERIAL a ser procurado.
 * @param {string} itemComponente - O ITEM_COMPONENTE a ser procurado.
 * @param {HTMLTableRowElement} [currentRow=null] - A linha atual que está sendo editada (para ignorar na busca).
 * @returns {HTMLTableRowElement|null} A primeira linha duplicada encontrada, ou null se nenhuma for encontrada.
 */
function encontrarLinhaDuplicada(codigoMaterial, itemComponente, currentRow = null) {
    if (!codigoMaterial || !itemComponente) return null;

    const linhas = Array.from(tabela.rows);
    for (const row of linhas) {
        if (currentRow && row === currentRow) continue;

        const data = getLinhaData(row);
        if (data.CODIGO_MATERIAL.toUpperCase() === codigoMaterial.toUpperCase() &&
            data.ITEM_COMPONENTE.toUpperCase() === itemComponente.toUpperCase()) {
            return row;
        }
    }
    return null;
}

/**
 * Exibe um SweetAlert2 com opções para lidar com uma duplicata encontrada.
 * @param {Object} newData - Os dados da nova linha ou linha que está sendo colada/digitada.
 * @param {HTMLTableRowElement} existingRow - A linha existente que é uma duplicata.
 * @returns {Promise<string>} Uma promessa que resolve com a ação escolhida pelo usuário ('ignorar', 'cancelar').
 */
async function mostrarAlertaDuplicata(newData, existingRow) {
    const existingData = getLinhaData(existingRow);
    const qtdeNova = parseFloat(String(newData.QTDE_MONTAGEM).replace(',', '.')) || 0;
    const qtdeExistente = parseFloat(String(existingData.QTDE_MONTAGEM).replace(',', '.')) || 0;

    const qtdeNovaFormatada = qtdeNova.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    const qtdeExistenteFormatada = qtdeExistente.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 });

    const result = await Swal.fire({
        title: '⚠️ Duplicata Encontrada!',
        html: `
            <p>A combinação de <strong>Código Material: ${newData.CODIGO_MATERIAL}</strong> e <strong>Item Componente: ${newData.ITEM_COMPONENTE}</strong> já existe na linha <strong>${Array.from(tabela.rows).indexOf(existingRow) + 1}</strong>.</p>
            <p><strong>Quantidade atual na linha existente:</strong> ${qtdeExistenteFormatada}<br>
            <strong>Quantidade na linha a ser inserida:</strong> ${qtdeNovaFormatada}</p>
        `,
        icon: 'warning',
        showCancelButton: true, // Corresponde ao botão "Cancel"
        confirmButtonText: 'Ignorar e Inserir', // Este é agora o botão 'confirm'
        showDenyButton: false, // Desabilita o botão "Não Adicionar"
        allowOutsideClick: false,
        allowEscapeKey: false,
        reverseButtons: true // Mantém a ordem dos botões como no layout anterior
    });

    if (result.isConfirmed) { // "Ignorar e Inserir"
        return 'ignorar';
    } else if (result.dismiss === Swal.DismissReason.cancel) { // "Cancel"
        return 'cancelar';
    }
    return 'cancelar'; // Default para fechar fora ou por outras razões
}

// --- FIM DAS NOVAS FUNÇÕES PARA DUPLICATAS ---


/**
 * Cria um elemento <td> contendo um <input> ou <select>.
 * Aplica event listeners para input, navegação com Enter, e demarcação/pintura.
 * @param {string} type - O tipo do input (e.g., "text").
 * @param {boolean} [readOnly=false] - Se o input deve ser somente leitura.
 * @param {string} [value=""] - O valor inicial do input.
 * @param {boolean} [isPasteTarget=false] - Indica se a célula pode ser alvo de colagem.
 * @param {string} [className=""] - Classes CSS adicionais para o <td>.
 * @returns {HTMLTableCellElement} O elemento <td> criado.
 */
function inputCell(type, readOnly = false, value = "", isPasteTarget = false, className = "") {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = type;
    input.readOnly = readOnly;
    input.value = (value || "");

    if (className) td.classList.add(className);

    // Event listener para mudanças de input
    input.addEventListener("input", async (e) => {
        // Aplica a formatação (maiúsculas/minúsculas)
        if (e.target.closest('td').classList.contains('unidade-medida-col')) {
            e.target.value = e.target.value.toLowerCase();
        } else {
            e.target.value = e.target.value.toUpperCase();
        }

        const currentRow = e.target.closest('tr');
        const currentData = getLinhaData(currentRow);

        const isCodigoMaterialCol = e.target.closest('td').classList.contains('codigo-material-col');
        const isItemComponenteCol = e.target.closest('td').classList.contains('item-componente-col');

        // Dispara o alerta de duplicata se CODIGO_MATERIAL e ITEM_COMPONENTE estiverem preenchidos
        // E SE a mudança ocorreu em CODIGO_MATERIAL ou ITEM_COMPONENTE
        if (currentData.CODIGO_MATERIAL && currentData.ITEM_COMPONENTE &&
            (isCodigoMaterialCol || isItemComponenteCol)) {

            const existingDuplicateRow = encontrarLinhaDuplicada(
                currentData.CODIGO_MATERIAL,
                currentData.ITEM_COMPONENTE,
                currentRow
            );

            if (existingDuplicateRow) {
                const action = await mostrarAlertaDuplicata(currentData, existingDuplicateRow);

                // --- Remova quaisquer efeitos de opacidade/destaque temporário após a escolha do SweetAlert ---
                resetDuplicateSearchState(); // Chama a função para limpar o estado de busca e destaques.

                if (action === 'ignorar') {
                    currentRow.classList.remove("no-highlight-on-ignore");
                    Swal.fire("ℹ️ Duplicata Ignorada", "A linha será inserida normalmente e destacada.", "info");
                } else if (action === 'cancelar') { // Seu "Não Adicionar" (agora botão Cancelar)
                    e.target.value = "";
                    const otherMaterialInput = currentRow.querySelector(".codigo-material-col input");
                    const otherItemInput = currentRow.querySelector(".item-componente-col input");
                    const qtdeInput = currentRow.querySelector(".qtde-montagem-col input");

                    if (otherMaterialInput && otherMaterialInput !== e.target) otherMaterialInput.value = "";
                    if (otherItemInput && otherItemInput !== e.target) otherItemInput.value = "";
                    if (qtdeInput) qtdeInput.value = ""; // Limpa quantidade também, se existir e não for o input atual
                    currentRow.classList.remove("highlight-duplicate");
                    currentRow.classList.add("no-highlight-on-ignore"); // Impede que ela seja roxa.
                    Swal.fire("❌ Entrada Cancelada", "Os campos foram limpos para evitar duplicata.", "info");
                }
            }
        }

        verificarDuplicatas();
        if (td.classList.contains('nivel-col')) {
            aplicarIndentacao(e.target.closest('tr'));
            e.target.dispatchEvent(new Event('change', { bubbles: true }));
        }
        atualizarColunaLinha();
    });

    // --- LISTENER: Desativa o foco e opacidade ao clicar em qualquer input/select de uma linha destacada ---
    input.addEventListener("click", (e) => {
        const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
        const clickedRow = e.target.closest('tr');

        // Verifica se o tbody está com o fade ativo OU se a linha clicada é uma linha focada/temporariamente destacada
        if (tbodyElement.classList.contains("table-faded") || clickedRow.classList.contains('highlight-focused-item')) {
            resetDuplicateSearchState(); // Limpa o estado de busca e destaques ao clicar.
        }
    });


    // Navegação com Enter (mantido)
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            const currentInput = e.target;
            const currentTd = currentInput.closest('td');
            const currentRow = currentInput.closest('tr');
            const currentRowIndex = Array.from(tabela.rows).indexOf(currentRow);
            const currentCellIndex = Array.from(currentRow.children).indexOf(currentTd);

            const nextRow = tabela.rows[currentRowIndex + 1];
            if (nextRow) {
                const nextTd = nextRow.children[currentCellIndex];
                const nextInput = nextTd?.querySelector('input, select');
                if (nextInput) {
                    nextInput.focus();
                } else {
                    const nextCellInRow = currentRow.children[currentCellIndex + 1];
                    const nextInputInRow = nextCellInRow?.querySelector('input, select');
                    if (nextInputInRow) {
                        nextInputInRow.focus();
                    }
                }
            } else {
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                acaoImportouOuAdicionouLinhas();
                const firstInputInNewRow = newRow.children[currentCellIndex]?.querySelector('input, select');
                if (firstInputInNewRow) {
                    firstInputInNewRow.focus();
                }
            }
        }
    });

    // Funcionalidade de pintura/demarcação de célula (mantida)
    input.addEventListener("click", (e) => {
        if (!demarcarLinha) { // Apenas se o modo "Demarcar linha" NÃO estiver ativo
            if (removerDemarcacao) {
                e.target.closest("td").style.backgroundColor = ""; // Remove cor
            } else if (corSelecionada) {
                // Alterna a cor: se já tem a cor selecionada, remove; senão, aplica
                if (e.target.closest("td").style.backgroundColor === corSelecionada) {
                    e.target.closest("td").style.backgroundColor = "";
                } else {
                    e.target.closest("td").style.backgroundColor = corSelecionada;
                }
            }
        }
    });

    td.appendChild(input);
    return td;
}

/**
 * Cria um elemento <td> contendo um <select>.
 * Aplica event listeners para mudança de valor, navegação com Enter, e demarcação/pintura.
 * @param {Array<string>} [options=[]] - As opções para o <select>.
 * @param {string} [selected=""] - A opção pré-selecionada.
 * @param {string} [className=""] - Classes CSS adicionais para o <td>.
 * @returns {HTMLTableCellElement} O elemento <td> criado.
 */
function selectCell(options = [], selected = "", className = "") {
    const td = document.createElement("td");
    const select = document.createElement("select");
    if (className) td.classList.add(className);

    options.forEach(opt => {
        const option = document.createElement("option");
        option.value = opt;
        option.textContent = opt;
        if (opt === selected) option.selected = true;
        select.appendChild(option);
    });

    // Event listener para mudanças no select
    select.addEventListener("change", () => {
        verificarDuplicatas(); // Verifica duplicatas a cada alteração
        atualizarColunaLinha(); // Sempre atualiza LINHA se campos relevantes mudam
    });

    // Funcionalidade de pintura/demarcação de célula (mantida)
    select.addEventListener("click", (e) => {
        if (!demarcarLinha) { // Apenas se o modo "Demarcar linha" NÃO estiver ativo
            if (removerDemarcacao) {
                e.target.closest("td").style.backgroundColor = "";
            } else if (corSelecionada) {
                if (e.target.closest("td").style.backgroundColor === corSelecionada) {
                    e.target.closest("td").style.backgroundColor = "";
                } else {
                    e.target.closest("td").style.backgroundColor = corSelecionada;
                }
            }
        }
    });

    // Navegação com Enter (mantido)
    select.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            const currentSelect = e.target;
            const currentTd = currentSelect.closest('td');
            const currentRow = currentSelect.closest('tr');
            const currentRowIndex = Array.from(tabela.rows).indexOf(currentRow);
            const currentCellIndex = Array.from(currentRow.children).indexOf(currentTd);

            const nextRow = tabela.rows[currentRowIndex + 1];
            if (nextRow) {
                const nextTd = nextRow.children[currentCellIndex];
                const nextInput = nextTd?.querySelector('input, select');
                if (nextInput) {
                    nextInput.focus();
                } else {
                    const nextCellInRow = currentRow.children[currentCellIndex + 1];
                    const nextInputInRow = nextCellInRow?.querySelector('input, select');
                    if (nextInputInRow) {
                        nextInputInRow.focus();
                    }
                }
            } else {
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                acaoImportouOuAdicionouLinhas();
                const firstInputInNewRow = newRow.children[currentCellIndex]?.querySelector('input, select');
                if (firstInputInNewRow) {
                    firstInputInNewRow.focus();
                }
            }
        }
    });

    td.appendChild(select);
    return td;
}

/**
 * Cria uma nova linha (<tr>) na tabela com todas as células padrão (input, select, checkbox).
 * @param {Object} [v={}] - Um objeto contendo os valores iniciais para preencher as células.
 * @returns {HTMLTableRowElement} A linha (<tr>) criada.
 */
function criarLinha(v = {}) {
    const row = document.createElement("tr");

    // Célula do checkbox de seleção
    const checkboxTd = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.classList.add("linha-selecao");
    checkboxTd.appendChild(checkbox);
    row.appendChild(checkboxTd);

    // Célula para a sequência (SEQ) - preenchida via JS
    const seqTd = document.createElement("td");
    seqTd.classList.add("seq-col");
    row.appendChild(seqTd);

    // Célula para NÍVEL
    const nivelCell = inputCell("text", false, v.NIVEL || "", true, "nivel-col");
    row.appendChild(nivelCell);

    // Event listener para demarcação/pintura de LINHA inteira
    row.addEventListener("click", (e) => {
        if (e.target.tagName.match(/INPUT|SELECT|BUTTON/)) return;
        if (demarcarLinha) {
            if (removerDemarcacao) {
                row.style.backgroundColor = "";
            } else if (corSelecionada) {
                if (rgbToHex(row.style.backgroundColor) === corSelecionada.toUpperCase()) {
                    row.style.backgroundColor = "";
                } else {
                    row.style.backgroundColor = corSelecionada;
                }
            }
        }
    });

    // Adiciona as demais células utilizando as funções helper
    row.appendChild(selectCell(siteValores, v.SITE || "1"));
    row.appendChild(selectCell(alternativas, v.ALTERNATIVA || "*"));
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || "", true, "codigo-material-col"));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));

    const linhaCell = inputCell("text", true, v.LINHA || "");
    linhaCell.classList.add("linha-auto-col");
    row.appendChild(linhaCell);

    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || "", true, "item-componente-col"));
    row.appendChild(inputCell("text", false, v.QTDE_MONTAGEM || "", false, "qtde-montagem-col"));
    row.appendChild(inputCell("text", false, (v.UNIDADE_MEDIDA || "").toLowerCase(), true, "unidade-medida-col"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    aplicarIndentacao(row);

    if (!seqAtivo) seqTd.style.display = "none";
    if (!nivelColVisivel) nivelCell.style.display = "none";

    return row;
}

/**
 * Cria uma nova linha vazia, usando a função `criarLinha` sem valores iniciais.
 * @returns {HTMLTableRowElement} A linha vazia criada.
 */
function criarLinhaVazia() {
    return criarLinha({});
}

/**
 * Adiciona 10 novas linhas vazias à tabela.
 */
function criar10Linhas() {
    for (let i = 0; i < 10; i++) {
        tabela.appendChild(criarLinhaVazia());
    }
}

/**
 * Atualiza os números de sequência (coluna "SEQ") para todas as linhas da tabela.
 * Cada linha recebe um número múltiplo de 10 (10, 20, 30...).
 */
function atualizarSequencias() {
    const linhas = tabela.querySelectorAll("tr");
    linhas.forEach((row, index) => {
        const seqTd = row.querySelectorAll("td")[1];
        if (seqTd) seqTd.textContent = (index + 1) * 1;
    });
}

/**
 * Preenche automaticamente a coluna "LINHA" baseando-se no `CODIGO_MATERIAL` e `ITEM_COMPONENTE`.
 *
 * Regras:
 * 1. Se o `ITEM_COMPONENTE` de UMA LINHA começar com "MP1-", APENAS ESSA LINHA terá a "LINHA" definida como "10".
 * Esta regra tem prioridade máxima. O sequenciamento das demais linhas (mesmo do mesmo CODIGO_MATERIAL)
 * NÃO É AFETADO ou reiniciado por uma linha "MP1-".
 * 2. Para TODAS AS OUTRAS LINHAS (cujo ITEM_COMPONENTE NÃO começa com "MP1-"):
 * a. Se um grupo de `CODIGO_MATERIAL` contiver qualquer `ITEM_COMPONENTE` que comece com "MP-" (E NÃO "MP1-"),
 * todas as linhas desse grupo (exceto as "MP1-") terão a "LINHA" definida como "10".
 * b. Caso contrário, a "LINHA" será uma sequência incremental (10, 20, 30...) para aquele `CODIGO_MATERIAL`.
 * O sequenciamento continua de onde parou para o mesmo CODIGO_MATERIAL, ou reinicia em 10 se o CODIGO_MATERIAL mudar.
 */
function atualizarColunaLinha() {
    const rows = Array.from(tabela.rows);
    const groupedData = new Map();

    rows.forEach(row => {
        const data = getLinhaData(row);
        const codigoMaterial = data.CODIGO_MATERIAL.trim();
        const itemComponente = data.ITEM_COMPONENTE.trim();

        if (codigoMaterial === "" && itemComponente === "") {
            const linhaInput = row.querySelectorAll("td")[7]?.querySelector("input");
            if (linhaInput) linhaInput.value = "";
            return;
        }

        if (itemComponente.toUpperCase().startsWith("MP1-")) {
            return;
        }

        if (codigoMaterial !== "") {
            if (!groupedData.has(codigoMaterial)) {
                groupedData.set(codigoMaterial, {
                    rows: [],
                    hasMPGeneral: false
                });
            }
            const group = groupedData.get(codigoMaterial);
            group.rows.push(row);
            if (itemComponente.toUpperCase().startsWith("MP-")) {
                group.hasMPGeneral = true;
            }
        }
    });

    let currentCodigoMaterial = "";
    let currentSequence = 10;

    rows.forEach(row => {
        const data = getLinhaData(row);
        const codigoMaterial = data.CODIGO_MATERIAL.trim();
        const itemComponente = data.ITEM_COMPONENTE.trim();
        const linhaInput = row.querySelectorAll("td")[7]?.querySelector("input");

        if (!linhaInput) return;

        if (codigoMaterial === "" && itemComponente === "") {
            currentCodigoMaterial = "";
            currentSequence = 10;
            return;
        }

        if (itemComponente.toUpperCase().startsWith("MP1-")) {
            linhaInput.value = "10";
            return;
        }

        if (codigoMaterial !== "") {
            if (codigoMaterial !== currentCodigoMaterial) {
                currentCodigoMaterial = codigoMaterial;
                currentSequence = 10;
            }

            const group = groupedData.get(codigoMaterial);
            if (group && group.hasMPGeneral) {
                linhaInput.value = "10";
            } else {
                linhaInput.value = String(currentSequence);
                currentSequence += 10;
            }
        } else {
            linhaInput.value = "";
            currentSequence = 10;
            currentCodigoMaterial = "";
        }
    });
}


/**
 * Aplica classes CSS à linha para indentação visual baseada no valor da coluna "NÍVEL".
 * @param {HTMLTableRowElement} row - A linha (<tr>) a ser indentada.
 */
function aplicarIndentacao(row) {
    for (let i = 1; i <= 10; i++) row.classList.remove(`nivel-${i}`);
    const nivelInput = row.querySelectorAll("td")[2]?.querySelector("input");
    if (nivelInput) {
        let nivel = parseInt(nivelInput.value);
        if (!isNaN(nivel) && nivel >= 1 && nivel <= 10) {
            row.classList.add(`nivel-${nivel}`);
        }
    }
}

/**
 * Extrai os dados de uma linha da tabela e os retorna como um objeto.
 * @param {HTMLTableRowElement} tr - A linha (<tr>) da qual extrair os dados.
 * @returns {Object} Um objeto com os dados da linha.
 */
function getLinhaData(tr) {
    const cells = tr.querySelectorAll("td");
    return {
        NIVEL: cells[2]?.querySelector("input")?.value.trim() || "",
        SITE: cells[3]?.querySelector("select")?.value || "",
        ALTERNATIVA: cells[4]?.querySelector("select")?.value || "",
        CODIGO_MATERIAL: cells[5]?.querySelector("input")?.value.trim().toUpperCase() || "",
        TIPO_ESTRUTURA: cells[6]?.querySelector("select")?.value || "",
        LINHA: cells[7]?.querySelector("input")?.value || "",
        ITEM_COMPONENTE: cells[8]?.querySelector("input")?.value.trim().toUpperCase() || "",
        QTDE_MONTAGEM: cells[9]?.querySelector("input")?.value.trim() || "",
        UNIDADE_MEDIDA: cells[10]?.querySelector("input")?.value.trim().toLowerCase() || "",
        FATOR_SUCATA: cells[11]?.querySelector("select")?.value || ""
    };
}

/**
 * Preenche as células de uma linha da tabela com os dados fornecidos.
 * @param {HTMLTableRowElement} row - A linha (<tr>) a ser preenchida.
 * @param {Object} data - Um objeto com os dados para preencher a linha.
 */
function preencherLinha(row, data) {
    const cells = row.querySelectorAll("td");
    cells[2].querySelector("input").value = data.NIVEL || "";
    aplicarIndentacao(row);
    cells[3].querySelector("select").value = data.SITE || "1";
    cells[4].querySelector("select").value = data.ALTERNATIVA || "*";
    cells[5].querySelector("input").value = (data.CODIGO_MATERIAL || "").toUpperCase();
    cells[6].querySelector("select").value = data.TIPO_ESTRUTURA || "Manufatura";
    cells[8].querySelector("input").value = (data.ITEM_COMPONENTE || "").toUpperCase();
    cells[9].querySelector("input").value = (data.QTDE_MONTAGEM === "0" ? "" : String(data.QTDE_MONTAGEM) || "").replace(",", ".");
    cells[10].querySelector("input").value = (data.UNIDADE_MEDIDA || "").toLowerCase();
    cells[11].querySelector("select").value = data.FATOR_SUCATA || "0";
}

/**
 * Verifica e destaca linhas que contêm dados duplicados na tabela.
 * Atualiza o contador de duplicatas no header.
 * Ignora linhas completamente vazias e o destaque pode ser desativado via checkbox.
 * ATUALIZADO: Agora considera duplicata APENAS se CODIGO_MATERIAL e ITEM_COMPONENTE forem idênticos.
 */
function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);

    // Remove todas as classes de destaque e a classe temporária de "não destaque"
    linhas.forEach(row => {
        row.classList.remove("highlight-duplicate");
        row.classList.remove("no-highlight-on-ignore");
    });

    if (ignorarDuplicatas) {
        document.getElementById("duplicateCountDisplay").textContent = ""; // Limpa o contador se ignorar
        resetDuplicateSearchState(); // Garante que o modo de busca seja desativado
        return;
    }

    const combinaçõesDetectadas = new Map();
    const tempFoundDuplicates = []; // Usa um array temporário para construir a lista

    linhas.forEach((tr) => {
        const data = getLinhaData(tr);
        if (data.CODIGO_MATERIAL === "" || data.ITEM_COMPONENTE === "") {
            return;
        }

        const hash = `${data.CODIGO_MATERIAL.toUpperCase()}|${data.ITEM_COMPONENTE.toUpperCase()}`;

        if (!combinaçõesDetectadas.has(hash)) {
            combinaçõesDetectadas.set(hash, []);
        }
        combinaçõesDetectadas.get(hash).push(tr);
    });

    let duplicateCount = 0;
    for (const [hash, rows] of combinaçõesDetectadas) {
        if (rows.length > 1) { // Se há mais de uma linha com essa combinação, são duplicatas
            rows.forEach(row => {
                if (!row.classList.contains("no-highlight-on-ignore")) {
                    row.classList.add("highlight-duplicate");
                }
            });
            // Adiciona todas as linhas duplicadas ao array temporário
            tempFoundDuplicates.push(...rows);
            duplicateCount += rows.length; // Conta cada ocorrência duplicada
        }
    }
    // Garante que foundDuplicates contenha apenas as duplicatas ativas (roxo claro)
    foundDuplicates = tempFoundDuplicates.filter(row => row.classList.contains("highlight-duplicate"));
    foundDuplicates.sort((a, b) => Array.from(tabela.rows).indexOf(a) - Array.from(tabela.rows).indexOf(b)); // Ordena por posição na tabela

    const displayElement = document.getElementById("duplicateCountDisplay");
    if (foundDuplicates.length > 0) { // Se o length for 0, não há itens duplicados para buscar
        displayElement.textContent = `⚠️ ${foundDuplicates.length} duplicata(s)`;
    } else {
        displayElement.textContent = "";
    }

    currentDuplicateIndex = -1; // Reseta o índice ao re-verificar duplicatas
}

/**
 * Converte uma cor RGB (string "rgb(r, g, b)") para sua representação hexadecimal.
 * @param {string} rgb - A string da cor RGB.
 * @returns {string} A string da cor em formato hexadecimal (e.g., "#RRGGBB").
 */
function rgbToHex(rgb) {
    if (!rgb || rgb.indexOf('rgb') === -1) return rgb ? rgb.toUpperCase() : "";
    const parts = rgb.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
    if (!parts) return "";
    delete parts[0];
    for (let i = 1; i <= 3; i++) {
        parts[i] = parseInt(parts[i]).toString(16);
        if (parts[i].length === 1) parts[i] = "0" + parts[i];
    }
    return "#" + parts.join("").toUpperCase();
}

/**
 * Exporta os dados da tabela atual para um arquivo Excel (.xlsx).
 * Inclui apenas as colunas especificadas e ignora linhas "vazias" (separadores).
 */
function exportarParaExcel() {
    const ws_data = [
        ["NIVEL", "SITE", "ALTERNATIVA", "CODIGO_MATERIAL", "TIPO ESTRUTURA", "LINHA", "ITEM_COMPONENTE", "QTDE_MONTAGEM", "UNIDADE DE MEDIDA", "FATOR_SUCATA"]
    ];

    tabela.querySelectorAll("tr").forEach(row => {
        const rowData = getLinhaData(row);

        if (rowData.CODIGO_MATERIAL === "" && rowData.ITEM_COMPONENTE === "") {
            return;
        }

        const dataRow = [
            rowData.NIVEL,
            rowData.SITE,
            rowData.ALTERNATIVA,
            rowData.CODIGO_MATERIAL,
            rowData.TIPO_ESTRUTURA,
            rowData.LINHA,
            rowData.ITEM_COMPONENTE,
            rowData.QTDE_MONTAGEM,
            rowData.UNIDADE_MEDIDA,
            rowData.FATOR_SUCATA
        ];
        ws_data.push(dataRow);
    });

    if (ws_data.length <= 1) {
        Swal.fire("ℹ️ Nada para Exportar", "A tabela está vazia ou contém apenas linhas sem dados preenchidos.", "info");
        return;
    }

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lista Tecnica");

    const now = new Date();
    const dateStr = now.getFullYear() + "-" +
                    String(now.getMonth() + 1).padStart(2, '0') + "-" +
                    String(now.getDate()).padStart(2, '0') + "_" +
                    String(now.getHours()).padStart(2, '0') + "-" +
                    String(now.getMinutes()).padStart(2, '0') + "-" +
                    String(now.getSeconds()).padStart(2, '0');

    XLSX.writeFile(wb, `Lista_Tecnica_${dateStr}.xlsx`);

    Swal.fire("✅ Exportado!", `A lista foi exportada para 'Lista_Tecnica_${dateStr}.xlsx'.`, "success");

}

/**
 * Carrega dados de um arquivo Excel selecionado pelo usuário para a tabela.
 * Mapeia as colunas do Excel dinamicamente pelos seus cabeçalhos.
 * @param {HTMLInputElement} inputElement - O elemento <input type="file"> que disparou o evento.
 */

function carregarExcel(inputElement) {
    const file = inputElement.files[0];
    if (!file) {
        Swal.fire("⚠️ Nenhum arquivo selecionado", "Por favor, selecione um arquivo Excel.", "warning");
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (json.length === 0 || !json[0]) {
            Swal.fire("⚠️ Arquivo Vazio ou Inválido", "O arquivo Excel está vazio, não contém cabeçalhos ou dados.", "warning");
            return;
        }

        tabela.innerHTML = "";

        const headers = json[0].map(h => String(h).trim().replace(/\s/g, '_').toUpperCase());
        const dataRows = json.slice(1);

        const colIndices = {
            NIVEL: headers.indexOf("NIVEL"),
            SITE: headers.indexOf("SITE"),
            ALTERNATIVA: headers.indexOf("ALTERNATIVA"),
            CODIGO_MATERIAL: headers.indexOf("CODIGO_MATERIAL"),
            TIPO_ESTRUTURA: headers.indexOf("TIPO_ESTRUTURA"),
            LINHA: headers.indexOf("LINHA"),
            ITEM_COMPONENTE: headers.indexOf("ITEM_COMPONENTE"),
            QTDE_MONTAGEM: headers.indexOf("QTDE_MONTAGEM"),
            UNIDADE_MEDIDA: headers.indexOf("UNIDADE_DE_MEDIDA") !== -1 ?
                            headers.indexOf("UNIDADE_DE_MEDIDA") :
                            headers.indexOf("UNIDADE_MEDIDA"),
            FATOR_SUCATA: headers.indexOf("FATOR_SUCATA")
        };

        console.log("Cabeçalhos do Excel (normalizados):", headers);
        console.log("Índices de Colunas Mapeados:", colIndices);
        for (const key in colIndices) {
            if (colIndices[key] === -1) {
                console.warn(`Atenção: A coluna "${key}" não foi encontrada no Excel. Verifique o nome do cabeçalho.`);
            }
        }

        dataRows.forEach(rowData => {
            const rowObj = {
                NIVEL: colIndices.NIVEL !== -1 && rowData[colIndices.NIVEL] !== undefined ? String(rowData[colIndices.NIVEL]) : "",
                SITE: colIndices.SITE !== -1 && rowData[colIndices.SITE] !== undefined ? String(rowData[colIndices.SITE]) : "1",
                ALTERNATIVA: colIndices.ALTERNATIVA !== -1 && rowData[colIndices.ALTERNATIVA] !== undefined ? String(rowData[colIndices.ALTERNATIVA]) : "*",
                CODIGO_MATERIAL: colIndices.CODIGO_MATERIAL !== -1 && rowData[colIndices.CODIGO_MATERIAL] !== undefined ? String(rowData[colIndices.CODIGO_MATERIAL]) : "",
                TIPO_ESTRUTURA: colIndices.TIPO_ESTRUTURA !== -1 && rowData[colIndices.TIPO_ESTRUTURA] !== undefined ? String(rowData[colIndices.TIPO_ESTRUTURA]) : "Manufatura",
                ITEM_COMPONENTE: colIndices.ITEM_COMPONENTE !== -1 && rowData[colIndices.ITEM_COMPONENTE] !== undefined ? String(rowData[colIndices.ITEM_COMPONENTE]) : "",
                QTDE_MONTAGEM: colIndices.QTDE_MONTAGEM !== -1 && rowData[colIndices.QTDE_MONTAGEM] !== undefined ? String(rowData[colIndices.QTDE_MONTAGEM]) : "",
                UNIDADE_MEDIDA: colIndices.UNIDADE_MEDIDA !== -1 && rowData[colIndices.UNIDADE_MEDIDA] !== undefined ? String(rowData[colIndices.UNIDADE_MEDIDA]).toLowerCase() : "",
                FATOR_SUCATA: colIndices.FATOR_SUCATA !== -1 && rowData[colIndices.FATOR_SUCATA] !== undefined ? String(rowData[colIndices.FATOR_SUCATA]) : "0"
            };
            const newRow = criarLinha(rowObj);
            tabela.appendChild(newRow);
        });

        acaoImportouOuAdicionouLinhas();
        Swal.fire("✅ Importado!", "Os dados do Excel foram carregados e atualizados.", "success");
    };
    reader.readAsArrayBuffer(file);
}

/**
 * Função unificada para disparar todas as atualizações necessárias
 * após ações que modificam o conteúdo ou a estrutura das linhas da tabela.
 * Inclui: atualização de sequências, autopreenchimento de LINHA, e verificação de duplicatas.
 */
function acaoImportouOuAdicionouLinhas() {
    atualizarSequencias();
    atualizarColunaLinha();
    verificarDuplicatas();
}

// --- Event Listeners para Botões da Barra de Ferramentas ---

/**
 * Listener para o botão "Criar Nova Lista".
 * Limpa a tabela e adiciona 10 linhas vazias.
 */
document.getElementById("criarListaBtn").addEventListener("click", () => {
    tabela.innerHTML = "";
    criar10Linhas();
    acaoImportouOuAdicionouLinhas();
    Swal.fire("✅ Lista Criada!", "10 novas linhas foram adicionadas.", "success");
});

/**
 * Listener para o botão "Continuar Lista".
 * Adiciona 10 linhas vazias ao final da tabela existente.
 */
document.getElementById("continuarListaBtn").addEventListener("click", () => {
    criar10Linhas();
    acaoImportouOuAdicionouLinhas();
    Swal.fire("➕ Adicionado", "10 novas linhas foram inseridas.", "success");
});

/**
 * Listener para o botão "Salvar Lista".
 * Chama a função para exportar a tabela para Excel.
 */
document.getElementById("salvarListaBtn").addEventListener("click", exportarParaExcel);

/**
 * Listener para o botão "Copiar Selecionado".
 * Copia os dados das linhas selecionadas para um cache.
 */
document.getElementById("copiarSelecionadoBtn").addEventListener("click", () => {
    const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
    if (linhasSelecionadas.length === 0) {
        Swal.fire("⚠️ Nada para Copiar", "Nenhuma linha selecionada para cópia.", "warning");
        return;
    }
    cacheCopiado = linhasSelecionadas.map(row => getLinhaData(row));
    Swal.fire("✅ Copiado!", `${cacheCopiado.length} linhas copiadas.`, "success");
});

/**
 * Listener para o botão "Colar".
 * Cola os dados do cache em novas linhas ou sobrescreve linhas selecionadas.
 * ATUALIZADO: NÃO USA A NOVA LÓGICA DE DUPLICATAS COM SPLASH.
 * A lógica de DUPLICATAS COM SPLASH é implementada no `handlePasteMultipleLines` para colagem de texto em células.
 * Este botão "Colar" é para colar LINHAS COPIADAS (de `cacheCopiado`).
 */
document.getElementById("colarBtn").addEventListener("click", () => {
    if (cacheCopiado.length === 0) {
        Swal.fire("ℹ️ Nada para Colar", "Nenhum dado copiado.", "info");
        return;
    }

    const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
    const startIndex = linhasSelecionadas.length > 0 ? Array.from(tabela.rows).indexOf(linhasSelecionadas[0]) : tabela.rows.length;

    cacheCopiado.forEach((rowData, i) => {
        const targetRow = tabela.rows[startIndex + i];
        if (targetRow) {
            preencherLinha(targetRow, rowData);
        } else {
            const newRow = criarLinha(rowData);
            tabela.appendChild(newRow);
        }
    });

    acaoImportouOuAdicionouLinhas();
    Swal.fire("✅ Colado!", `${cacheCopiado.length} linhas coladas.`, "success");
});

/**
 * Listener para o botão "Deletar Selecionados".
 * Remove as linhas da tabela que estão marcadas com o checkbox.
 */
document.getElementById("deletarSelecionadosBtn").addEventListener("click", () => {
    const linhasParaDeletar = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
    if (linhasParaDeletar.length === 0) {
        Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione as linhas para deletar.", "warning");
        return;
    }

    Swal.fire({
        title: 'Tem certeza?',
        text: `Você vai deletar ${linhasParaDeletar.length} linha(s).`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d33',
        cancelButtonColor: '#3085d6',
        confirmButtonText: 'Sim, deletar!',
        cancelButtonText: 'Cancelar'
    }).then((result) => {
        if (result.isConfirmed) {
            linhasParaDeletar.forEach(row => row.remove()); // Remove as linhas
            acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas
            Swal.fire('Deletado!', `${linhasParaDeletar.length} linha(s) foram deletadas.`, 'success');
        }
    });
});

/**
 * Listener para o botão "Inserir Acima".
 * Insere uma nova linha vazia acima da primeira linha selecionada.
 */
document.getElementById("inserirAcimaBtn").addEventListener("click", () => {
    const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
    if (linhasSelecionadas.length === 0) {
        Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione a(s) linha(s) acima da qual deseja inserir.", "warning");
        return;
    }
    const primeiraLinhaSelecionada = linhasSelecionadas[0];
    const novaLinha = criarLinhaVazia();
    tabela.insertBefore(novaLinha, primeiraLinhaSelecionada);
    acaoImportouOuAdicionouLinhas();
    Swal.fire("⬆️ Inserido", "Uma nova linha foi inserida acima da seleção.", "success");
});

/**
 * Listener para o botão "Inserir Abaixo".
 * Insere uma nova linha vazia abaixo da última linha selecionada.
 */
document.getElementById("inserirAbaixoBtn").addEventListener("click", () => {
    const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
    if (linhasSelecionadas.length === 0) {
        Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione a(s) linha(s) abaixo da qual deseja inserir.", "warning");
        return;
    }
    const ultimaLinhaSelecionada = linhasSelecionadas[linhasSelecionadas.length - 1];
    const novaLinha = criarLinhaVazia();
    if (ultimaLinhaSelecionada.nextElementSibling) {
        tabela.insertBefore(novaLinha, ultimaLinhaSelecionada.nextElementSibling);
    } else {
        tabela.appendChild(novaLinha);
    }
    acaoImportouOuAdicionouLinhas();
    Swal.fire("⬇️ Inserido", "Uma nova linha foi inserida abaixo da seleção.", "success");
});

/**
 * Listener para o botão "Alternar SEQ".
 * Alterna a visibilidade da coluna "SEQ".
 */
document.getElementById("toggleSeqBtn").addEventListener("click", () => {
    seqAtivo = !seqAtivo;
    document.getElementById("listaTabela").classList.toggle("seq-col-hidden", !seqAtivo);
    document.getElementById("toggleSeqBtn").innerHTML = seqAtivo ? '<span class="material-symbols-outlined">visibility</span> SEQ' : '<span class="material-symbols-outlined">visibility_off</span> SEQ';
});

/**
 * Listener para o botão "Alternar NÍVEL".
 * Alterna a visibilidade da coluna "NÍVEL".
 */
document.getElementById("toggleNivelColBtn").addEventListener("click", () => {
    nivelColVisivel = !nivelColVisivel;
    document.getElementById("listaTabela").classList.toggle("nivel-col-hidden", !nivelColVisivel);
    document.getElementById("toggleNivelColBtn").innerHTML = nivelColVisivel ? '<span class="material-symbols-outlined">layers</span> NÍVEL' : '<span class="material-symbols-outlined">layers_clear</span> NÍVEL';
});

/**
 * Listener para o botão "Alternar Régua".
 * Alterna o efeito de destaque de linha ao passar o mouse (régua).
 */
document.getElementById("toggleHoverEffectBtn").addEventListener("click", () => {
    hoverEffectAtivo = !hoverEffectAtivo;
    const tableElement = document.getElementById("listaTabela");
    tableElement.classList.toggle("hover-effect", hoverEffectAtivo);
    tableElement.classList.toggle("no-hover-effect", !hoverEffectAtivo);
    document.getElementById("toggleHoverEffectBtn").innerHTML = hoverEffectAtivo ? '<span class="material-symbols-outlined">straighten</span> Régua' : '<span class="material-symbols-outlined">format_line_spacing</span> Régua';
});

// --- Event Listeners para Controles de Pintura e Destaque ---

/**
 * Listener para o botão "Limpar Seleção de Cor".
 * Reseta a cor selecionada e os modos de demarcação.
 */
document.getElementById("clearPaintBtn").addEventListener("click", () => {
    corSelecionada = "";
    demarcarLinha = false;
    removerDemarcacao = false;
    document.getElementById("demarcarLinhaCheckbox").checked = false;
    document.getElementById("removerDemarcacaoCheckbox").checked = false;
    Swal.fire("🎨 Limpeza", "Seleção de cor e modos de demarcação limpos.", "info");
});

/**
 * Listener para o checkbox "Demarcar linha".
 * Ativa/desativa o modo de demarcação de linha inteira.
 */
document.getElementById("demarcarLinhaCheckbox").addEventListener("change", (e) => {
    demarcarLinha = e.target.checked;
    if (demarcarLinha) removerDemarcacao = false;
    document.getElementById("removerDemarcacaoCheckbox").checked = false;
});

/**
 * Listener para o checkbox "Remover demarcação".
 * Ativa/desativa o modo de remoção de cor ao clicar.
 */
document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", (e) => {
    removerDemarcacao = e.target.checked;
    if (removerDemarcacao) demarcarLinha = false;
    document.getElementById("demarcarLinhaCheckbox").checked = false;
});

// --- Geração Dinâmica de Botões de Cor de Nível ---

/**
 * Cria dinamicamente botões para seleção de cores de Nível.
 * Cada botão define `corSelecionada` para a cor correspondente.
 */
const nivelColorButtonsDiv = document.getElementById("nivelColorButtons");
nivelColors.forEach((color, index) => {
    const button = document.createElement("button");
    button.classList.add("paint-btn");
    button.style.backgroundColor = color;
    button.style.color = getContrastColor(color);
    button.textContent = `Nível ${index + 1}`;
    button.dataset.color = color;
    button.addEventListener("click", (e) => {
        corSelecionada = e.target.dataset.color;
        Swal.fire(`🎨 Cor Selecionada`, `Cor para Nível ${index + 1} selecionada.`, "info");
    });
    nivelColorButtonsDiv.appendChild(button);
});

/**
 * Adiciona listeners aos botões de cores de atenção (predefinidos no HTML).
 * Define `corSelecionada` para a cor do botão clicado.
 */
document.querySelectorAll("#attentionColorButtons .paint-btn").forEach(button => {
    button.addEventListener("click", (e) => {
        corSelecionada = e.target.dataset.color;
        Swal.fire(`🎨 Cor Selecionada`, `Cor ${e.target.textContent.trim()} selecionada.`, "info");
    });
});

/**
 * Função utilitária para determinar uma cor de texto de contraste (preto ou branco)
 * para um dado fundo hexadecimal, baseando-se no valor de luminância.
 * @param {string} hexcolor - A cor hexadecimal do fundo (e.g., "#RRGGBB").
 * @returns {string} "black" ou "white".
 */
function getContrastColor(hexcolor) {
    if (!hexcolor.startsWith("#")) {
        return "black";
    }
    const r = parseInt(hexcolor.slice(1, 3), 16);
    const g = parseInt(hexcolor.slice(3, 5), 16);
    const b = parseInt(hexcolor.slice(5, 7), 16);
    const hsp = Math.sqrt(
        0.299 * (r * r) +
        0.587 * (g * g) +
        0.114 * (b * b)
    );
    return (hsp > 127.5) ? "black" : "white";
}

/**
 * Listener para o checkbox "Ignorar destaque de duplicatas".
 * Ativa/desativa a funcionalidade de destaque de duplicatas e atualiza a tabela.
 */
document.getElementById("ignorarDuplicatasCheckbox").addEventListener("change", (e) => {
    ignorarDuplicatas = e.target.checked;
    verificarDuplicatas();
});

// --- Sincronização de Checkboxes de Seleção Total ---

/**
 * Listener para o checkbox "Marcar/Desmarcar Todos" no painel de controle.
 * Sincroniza o estado com o checkbox do cabeçalho da tabela e marca/desmarca todas as linhas.
 */
document.getElementById("toggleAllCheckboxesHeader").addEventListener("click", (e) => {
    const isChecked = e.target.checked;
    document.getElementById("toggleAllCheckboxes").checked = isChecked;
    tabela.querySelectorAll(".linha-selecao").forEach(checkbox => {
        checkbox.checked = isChecked;
    });
});

/**
 * Listener para o checkbox "Marcar/Desmarcar Todos" no cabeçalho da tabela.
 * Sincroniza o estado com o checkbox do painel de controle e marca/desmarca todas as linhas.
 */
document.getElementById("toggleAllCheckboxes").addEventListener("click", (e) => {
    const isChecked = e.target.checked;
    document.getElementById("toggleAllCheckboxesHeader").checked = isChecked;
    tabela.querySelectorAll(".linha-selecao").forEach(checkbox => {
        checkbox.checked = isChecked;
    });
});

// --- Listener para Input de Arquivo (Importação) ---

/**
 * Listener para o input de arquivo (botão "Importar Lista").
 * Dispara a função `carregarExcel` quando um arquivo é selecionado.
 */
document.getElementById("inputFile").addEventListener("change", function (e) {
    carregarExcel(e.target);
    e.target.value = '';
});


// --- Lógica de Busca e Navegação de Duplicatas por Botão e Teclado ---

document.getElementById("findDuplicatesBtn").addEventListener("click", () => {
    resetDuplicateSearchState(); // Limpa qualquer estado anterior

    // Filtra apenas as linhas que SÃO destacadas como duplicatas (as roxas)
    const currentDuplicates = Array.from(tabela.querySelectorAll('tr.highlight-duplicate'));

    if (currentDuplicates.length === 0) {
        Swal.fire("ℹ️ Sem Duplicatas", "Não há itens duplicados para buscar na lista.", "info");
        return;
    }

    // Ordena as duplicatas pela posição na tabela para garantir a ordem de navegação
    foundDuplicates = currentDuplicates.sort((a, b) => {
        return Array.from(tabela.rows).indexOf(a) - Array.from(tabela.rows).indexOf(b);
    });

    currentDuplicateIndex = 0; // Começa pelo primeiro item duplicado

    // Rola para o primeiro item e aplica os destaques
    focusOnDuplicate(foundDuplicates[currentDuplicateIndex]);
});

/**
 * Aplica os efeitos visuais de foco a uma linha duplicada.
 * @param {HTMLTableRowElement} rowToFocus - A linha duplicada a ser focada.
 */
function focusOnDuplicate(rowToFocus) {
    if (!rowToFocus) return; // Sai se a linha não existir

    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');

    // Remove o destaque de foco de TODAS as linhas antes de aplicar no novo item
    tabela.querySelectorAll('tr').forEach(row => {
        row.classList.remove('highlight-focused-item');
        row.classList.remove('temp-highlight-found');
    });

    // Aplica o fade ao tbody para escurecer as outras linhas
    tbodyElement.classList.add("table-faded");

    // Aplica destaque na linha focada
    rowToFocus.classList.add('highlight-focused-item');
    rowToFocus.classList.add('highlight-duplicate'); // Garante que continue roxa
    rowToFocus.classList.add('temp-highlight-found'); // Efeito de piscar

    rowToFocus.scrollIntoView({ behavior: 'smooth', block: 'center' }); // Rola suavemente
}

// --- Listener GLOBAL para Teclas ENTER e ESC (para navegação de duplicatas) ---
document.addEventListener('keydown', (e) => {
    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');

    // Só ativa a navegação por teclado se houver duplicatas encontradas e a tabela estiver em modo "faded"
    if (foundDuplicates.length > 0 && tbodyElement.classList.contains("table-faded")) {
        if (e.key === 'Enter') {
            e.preventDefault(); // Evita o comportamento padrão do Enter (como focar no próximo input)
            currentDuplicateIndex++;
            if (currentDuplicateIndex >= foundDuplicates.length) {
                currentDuplicateIndex = 0; // Volta para o início da lista se chegar ao final
            }
            focusOnDuplicate(foundDuplicates[currentDuplicateIndex]);
        } else if (e.key === 'Escape') {
            e.preventDefault(); // Evita o comportamento padrão do Esc
            resetDuplicateSearchState(); // Sai do modo de busca de duplicatas
            Swal.fire("✅ Busca Encerrada", "O modo de busca de duplicatas foi desativado.", "info");
        }
    }
});


// --- COLAGEM MASSA COM CONFIRMAÇÃO ---
// Adiciona o listener de paste ao document para pegar colagens em qualquer lugar
document.addEventListener('paste', handlePasteMultipleLines);

/**
 * Função para colar múltiplos valores em qualquer campo da tabela (com suporte para criar linhas e ignorar cabeçalho).
 * Ativada por um evento 'paste' no documento.
 * @param {ClipboardEvent} event - O evento de colagem.
 */
async function handlePasteMultipleLines(event) {
    const pastedText = (event.clipboardData || window.clipboardData).getData('text');
    if (!pastedText) return;

    const lines = pastedText
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line !== '');

    if (lines.length === 0) {
        return; // Não exibe Swal.fire, apenas retorna
    }

    const possibleHeaders = new Set([
        'seq', 'codigo_material', 'codigo', 'qtd', 'qtde', 'un', 'unidade_de_medida', 'unidade_medida', 'unidade', 'descricao', 'item_componente', 'item', 'linha', 'nivel', 'site', 'alternativa', 'tipo_estrutura', 'fator_sucata'
    ]);

    const normalizeHeader = str => str.normalize('NFD').replace(/[\u0300-\u036f\s_]/g, "").toLowerCase();
    const firstLineNormalized = normalizeHeader(lines[0]);

    const isHeader = Array.from(possibleHeaders).some(header => firstLineNormalized.startsWith(normalizeHeader(header)));

    const realLines = lines.slice(isHeader ? 1 : 0);

    if (realLines.length === 0) {
        return; // Não exibe Swal.fire, apenas retorna
    }

    const activeElement = document.activeElement;
    if (
        activeElement &&
        (activeElement.tagName === 'INPUT' || activeElement.tagName === 'SELECT') &&
        activeElement.closest('#listaTabela')
    ) {
        // Remove a confirmação inicial ao colar
        // const confirmPaste = await Swal.fire({...});
        // if (!confirmPaste.isConfirmed) { event.preventDefault(); return; }

        event.preventDefault(); // Previne a ação de colagem padrão do navegador

        const targetRow = activeElement.closest('tr');
        const targetTd = activeElement.closest('td');
        const rowIndex = Array.from(tabela.rows).indexOf(targetRow);
        const columnIndex = Array.from(targetRow.children).indexOf(targetTd);

        let itemsPastedCount = 0;

        for (let i = 0; i < realLines.length; i++) {
            let rowToProcess = tabela.rows[rowIndex + i];
            let isNewlyCreatedRow = false;

            if (!rowToProcess) {
                rowToProcess = criarLinhaVazia();
                tabela.appendChild(rowToProcess);
                isNewlyCreatedRow = true;
            }

            const inputToUpdate = rowToProcess.children[columnIndex]?.querySelector('input, select');
            if (!inputToUpdate) continue;

            let valueToPaste = realLines[i];
            let currentMaterial = "";
            let currentItemComponente = "";
            let currentQtdeMontagem = "0";

            // Aplica o valor colado ao campo correto, COM A FORMATAÇÃO (UPPERCASE/LOWERCASE)
            // Esta é a chave para evitar a duplicação do texto no input de material.
            if (targetTd.classList.contains('codigo-material-col')) {
                inputToUpdate.value = valueToPaste.toUpperCase();
            } else if (targetTd.classList.contains('item-componente-col')) {
                inputToUpdate.value = valueToPaste.toUpperCase();
            } else if (targetTd.classList.contains('qtde-montagem-col')) {
                inputToUpdate.value = valueToPaste.replace(',', '.');
            } else if (targetTd.classList.contains('unidade-medida-col')) {
                inputToUpdate.value = valueToPaste.toLowerCase();
            } else {
                inputToUpdate.value = valueToPaste.toUpperCase(); // Para outros campos de texto
            }

            // Dispara o evento 'input' para que outras lógicas (LINHA, etc.) sejam atualizadas
            inputToUpdate.dispatchEvent(new Event('input', { bubbles: true }));

            // Obter os dados da linha ATUALIZADOS APÓS A COLAGEM para verificação de duplicidade
            const updatedRowData = getLinhaData(rowToProcess);
            currentMaterial = updatedRowData.CODIGO_MATERIAL;
            currentItemComponente = updatedRowData.ITEM_COMPONENTE;
            currentQtdeMontagem = updatedRowData.QTDE_MONTAGEM;

            // Lógica de duplicata ao colar (sem splash, apenas soma se for o caso)
            if (currentMaterial && currentItemComponente) {
                const existingDuplicateRow = encontrarLinhaDuplicada(
                    currentMaterial,
                    currentItemComponente,
                    rowToProcess // Passa a própria linha para ignorá-la na busca por duplicatas
                );

                if (existingDuplicateRow) {
                    // Se for uma duplicata, soma a quantidade e remove a linha recém-colada.
                    const qtdeNova = parseFloat(currentQtdeMontagem) || 0;
                    const existingQtdeInput = existingDuplicateRow.querySelector(".qtde-montagem-col input");
                    if (existingQtdeInput) {
                        const currentQtde = parseFloat(existingQtdeInput.value.replace(',', '.')) || 0;
                        existingQtdeInput.value = (currentQtde + qtdeNova).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
                        existingQtdeInput.dispatchEvent(new Event('input', { bubbles: true })); // Dispara input na linha existente
                    }
                    if (rowToProcess.parentNode) {
                        rowToProcess.remove(); // Remove a linha recém-colada
                    }
                    itemsPastedCount++; // Conta como processado/somado
                    console.log(`[Colar] Duplicata encontrada para ${currentMaterial}|${currentItemComponente}. Quantidade somada à linha existente.`);
                } else {
                    // Não é duplicata, a linha permanece normalmente.
                    itemsPastedCount++;
                }
            } else {
                // Se não tem material e componente completos, apenas conta a linha colada.
                itemsPastedCount++;
            }
        }
        acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas após colar

        // Nenhuma mensagem final de sucesso/cancelamento. A operação é silenciosa.
        // Swal.fire("✅ Colagem concluída!", `Foram colados ${itemsPastedCount} item(s) com sucesso.`, "success");
    } else {
        console.log("Colagem em elemento não gerenciado pela tabela, permitindo padrão.");
    }
}


// --- Inicialização da Aplicação ---

/**
 * Executado quando o DOM está completamente carregado.
 * Inicializa a tabela com 10 linhas se estiver vazia, ou atualiza o conteúdo existente.
 * Define o estado inicial de visibilidade das colunas e efeitos.
 */
document.addEventListener("DOMContentLoaded", () => {
    if (tabela.rows.length === 0) {
        criar10Linhas();
        acaoImportouOuAdicionouLinhas();
        Swal.fire("🎉 Bem-vindo!", "A lista foi inicializada com 10 linhas para você começar.", "info");
    } else {
        acaoImportouOuAdicionouLinhas();
    }

    const tabelaElement = document.getElementById("listaTabela");
    if (!seqAtivo) tabelaElement.classList.add("seq-col-hidden");
    if (!nivelColVisivel) tabelaElement.classList.add("nivel-col-hidden");
    if (hoverEffectAtivo) tabelaElement.classList.add("hover-effect");
    else tabelaElement.classList.add("no-hover-effect");

    // Inicializa o texto dos botões de alternância de visibilidade com ícones Material Symbols
    document.getElementById("toggleSeqBtn").innerHTML = '<span class="material-symbols-outlined">visibility</span> SEQ';
    document.getElementById("toggleNivelColBtn").innerHTML = '<span class="material-symbols-outlined">layers</span> NÍVEL';
    document.getElementById("toggleHoverEffectBtn").innerHTML = '<span class="material-symbols-outlined">straighten</span> Régua';
});