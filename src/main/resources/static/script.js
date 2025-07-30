/**
 * @file script.js
 * @description Script principal para a aplicação de Lista Técnica IFS.
 * Gerencia todas as funcionalidades da interface, incluindo listeners de botões,
 * manipulação da tabela, tema escuro e alertas.
 * @version 4.0 - Versão estável com todas as funcionalidades originais e tema escuro.
 */

// Variáveis globais e de estado
let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true;
let nivelColVisivel = true;
let corSelecionada = "";
let demarcarLinha = false;
let removerDemarcacao = false;
let ignorarDuplicatas = false;
let hoverEffectAtivo = true;
const nivelColors = ["#4664cf", "#CD5C5C", "#B3E6B3", "#FFD700", "#8A2BE2", "#FF8C00", "#00CED1", "#FF69B4", "#9ACD32", "#DA70D6"];
const tiposEstrutura = ["Manufatura", "Comprado", ""];
const fatorSucata = ["0", "15", ""];
const alternativas = ["*", ""];
const siteValores = ["1", ""];
const niveis = Array.from({ length: 10 }, (_, i) => (i + 1).toString());
let foundDuplicates = [];
let currentDuplicateIndex = -1;

/**
 * Reseta o estado da busca e destaque de duplicatas na tabela.
 */
function resetDuplicateSearchState() {
    foundDuplicates = [];
    currentDuplicateIndex = -1;
    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
    tbodyElement.classList.remove("table-faded");
    tabela.querySelectorAll('tr.highlight-focused-item').forEach(row => {
        row.classList.remove('highlight-focused-item');
        row.classList.remove('temp-highlight-found');
    });
    verificarDuplicatas();
}

/**
 * Encontra a primeira linha na tabela que tem a mesma combinação de código material e item componente.
 * @param {string} codigoMaterial - O código do material a ser procurado.
 * @param {string} itemComponente - O item componente a ser procurado.
 * @param {HTMLTableRowElement|null} [currentRow=null] - A linha atual, para ser ignorada na busca.
 * @returns {HTMLTableRowElement|null} A linha duplicada encontrada, ou null se não houver.
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
 * Exibe um alerta SweetAlert quando uma duplicata é encontrada durante a digitação.
 * @param {object} newData - Os dados da linha atual que está sendo editada.
 * @param {HTMLTableRowElement} existingRow - A linha existente que é uma duplicata.
 * @returns {Promise<'ignorar'|'cancelar'>} A ação escolhida pelo usuário.
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
        showCancelButton: true,
        confirmButtonText: 'Ignorar e Inserir',
        showDenyButton: false,
        allowOutsideClick: false,
        allowEscapeKey: false,
        reverseButtons: true
    });

    if (result.isConfirmed) {
        return 'ignorar';
    } else if (result.dismiss === Swal.DismissReason.cancel) {
        return 'cancelar';
    }
    return 'cancelar';
}


/**
 * Cria e retorna um elemento de célula <td> com um <input> dentro.
 * @param {string} type - O tipo do input (ex: "text").
 * @param {boolean} [readOnly=false] - Define se o campo é somente leitura.
 * @param {string} [value=""] - O valor inicial do campo.
 * @param {boolean} [isPasteTarget=false] - Indica se é um alvo para colagem.
 * @param {string} [className=""] - Classes CSS para adicionar ao <td>.
 * @returns {HTMLTableCellElement} O elemento <td> criado.
 */
function inputCell(type, readOnly = false, value = "", isPasteTarget = false, className = "") {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = type;
    input.readOnly = readOnly;
    input.value = (value || "");

    if (className) td.classList.add(className);

    input.addEventListener("input", async (e) => {
        if (e.target.closest('td').classList.contains('unidade-medida-col')) {
            e.target.value = e.target.value.toLowerCase();
        } else {
            e.target.value = e.target.value.toUpperCase();
        }

        const currentRow = e.target.closest('tr');
        const currentData = getLinhaData(currentRow);
        const isCodigoMaterialCol = e.target.closest('td').classList.contains('codigo-material-col');
        const isItemComponenteCol = e.target.closest('td').classList.contains('item-componente-col');

        if (currentData.CODIGO_MATERIAL && currentData.ITEM_COMPONENTE &&
            (isCodigoMaterialCol || isItemComponenteCol)) {

            const existingDuplicateRow = encontrarLinhaDuplicada(
                currentData.CODIGO_MATERIAL,
                currentData.ITEM_COMPONENTE,
                currentRow
            );

            if (existingDuplicateRow) {
                const action = await mostrarAlertaDuplicata(currentData, existingDuplicateRow);
                resetDuplicateSearchState();

                if (action === 'ignorar') {
                    currentRow.classList.remove("no-highlight-on-ignore");
                    Swal.fire("ℹ️ Duplicata Ignorada", "A linha será inserida normalmente e destacada.", "info");
                } else if (action === 'cancelar') {
                    e.target.value = "";
                    const otherMaterialInput = currentRow.querySelector(".codigo-material-col input");
                    const otherItemInput = currentRow.querySelector(".item-componente-col input");
                    const qtdeInput = currentRow.querySelector(".qtde-montagem-col input");

                    if (otherMaterialInput && otherMaterialInput !== e.target) otherMaterialInput.value = "";
                    if (otherItemInput && otherItemInput !== e.target) otherItemInput.value = "";
                    if (qtdeInput) qtdeInput.value = "";
                    currentRow.classList.remove("highlight-duplicate");
                    currentRow.classList.add("no-highlight-on-ignore");
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

    input.addEventListener("click", (e) => {
        const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
        const clickedRow = e.target.closest('tr');

        if (tbodyElement.classList.contains("table-faded") || clickedRow.classList.contains('highlight-focused-item')) {
            resetDuplicateSearchState();
        }

        if (!demarcarLinha) {
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

    td.appendChild(input);
    return td;
}

/**
 * Cria e retorna um elemento de célula <td> com um <select> dentro.
 * @param {string[]} [options=[]] - Um array de strings para as opções do select.
 * @param {string} [selected=""] - O valor da opção que deve vir selecionada.
 * @param {string} [className=""] - Classes CSS para adicionar ao <td>.
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

    select.addEventListener("change", () => {
        verificarDuplicatas();
        atualizarColunaLinha();
    });

    select.addEventListener("click", (e) => {
        if (!demarcarLinha) {
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
 * Cria uma nova linha <tr> da tabela com todas as células e inputs.
 * @param {Object} [v={}] - Um objeto com os valores para preencher a linha.
 * @returns {HTMLTableRowElement} A linha <tr> criada.
 */
function criarLinha(v = {}) {
    const row = document.createElement("tr");

    const checkboxTd = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.classList.add("linha-selecao");
    checkboxTd.appendChild(checkbox);
    row.appendChild(checkboxTd);

    const seqTd = document.createElement("td");
    seqTd.classList.add("seq-col");
    row.appendChild(seqTd);

    const nivelCell = inputCell("text", false, v.NIVEL || "", true, "nivel-col");
    row.appendChild(nivelCell);

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

function criarLinhaVazia() { return criarLinha({}); }
function criar10Linhas() { for (let i = 0; i < 10; i++) { tabela.appendChild(criarLinhaVazia()); } }

function atualizarSequencias() {
    const linhas = tabela.querySelectorAll("tr");
    linhas.forEach((row, index) => {
        const seqTd = row.querySelectorAll("td")[1];
        if (seqTd) seqTd.textContent = (index + 1) * 1;
    });
}

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
        if (itemComponente.toUpperCase().startsWith("MP1-")) { return; }
        if (codigoMaterial !== "") {
            if (!groupedData.has(codigoMaterial)) {
                groupedData.set(codigoMaterial, { rows: [], hasMPGeneral: false });
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

function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    linhas.forEach(row => { row.classList.remove("highlight-duplicate"); row.classList.remove("no-highlight-on-ignore"); });
    if (ignorarDuplicatas) { document.getElementById("duplicateCountDisplay").textContent = ""; resetDuplicateSearchState(); return; }
    const combinaçõesDetectadas = new Map();
    const tempFoundDuplicates = [];
    linhas.forEach((tr) => {
        const data = getLinhaData(tr);
        if (data.CODIGO_MATERIAL === "" || data.ITEM_COMPONENTE === "") { return; }
        const hash = `${data.CODIGO_MATERIAL.toUpperCase()}|${data.ITEM_COMPONENTE.toUpperCase()}`;
        if (!combinaçõesDetectadas.has(hash)) { combinaçõesDetectadas.set(hash, []); }
        combinaçõesDetectadas.get(hash).push(tr);
    });
    let duplicateCount = 0;
    for (const [hash, rows] of combinaçõesDetectadas) {
        if (rows.length > 1) {
            rows.forEach(row => { if (!row.classList.contains("no-highlight-on-ignore")) { row.classList.add("highlight-duplicate"); } });
            tempFoundDuplicates.push(...rows);
            duplicateCount += rows.length;
        }
    }
    foundDuplicates = tempFoundDuplicates.filter(row => row.classList.contains("highlight-duplicate"));
    foundDuplicates.sort((a, b) => Array.from(tabela.rows).indexOf(a) - Array.from(tabela.rows).indexOf(b));
    const displayElement = document.getElementById("duplicateCountDisplay");
    if (foundDuplicates.length > 0) { displayElement.textContent = `⚠️ ${foundDuplicates.length} duplicata(s)`; } else { displayElement.textContent = ""; }
    currentDuplicateIndex = -1;
}

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

function focusOnDuplicate(rowToFocus) {
    if (!rowToFocus) return;
    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
    tabela.querySelectorAll('tr').forEach(row => { row.classList.remove('highlight-focused-item'); row.classList.remove('temp-highlight-found'); });
    tbodyElement.classList.add("table-faded");
    rowToFocus.classList.add('highlight-focused-item');
    rowToFocus.classList.add('highlight-duplicate');
    rowToFocus.classList.add('temp-highlight-found');
    rowToFocus.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function getContrastColor(hexcolor) {
    if (!hexcolor.startsWith("#")) { return "black"; }
    const r = parseInt(hexcolor.slice(1, 3), 16);
    const g = parseInt(hexcolor.slice(3, 5), 16);
    const b = parseInt(hexcolor.slice(5, 7), 16);
    const hsp = Math.sqrt(0.299 * (r * r) + 0.587 * (g * g) + 0.114 * (b * b));
    return (hsp > 127.5) ? "black" : "white";
}

function exportarParaExcel() {
    const ws_data = [["NIVEL", "SITE", "ALTERNATIVA", "CODIGO_MATERIAL", "TIPO ESTRUTURA", "LINHA", "ITEM_COMPONENTE", "QTDE_MONTAGEM", "UNIDADE DE MEDIDA", "FATOR_SUCATA"]];
    tabela.querySelectorAll("tr").forEach(row => {
        const rowData = getLinhaData(row);
        if (rowData.CODIGO_MATERIAL === "" && rowData.ITEM_COMPONENTE === "") { return; }
        const dataRow = [rowData.NIVEL, rowData.SITE, rowData.ALTERNATIVA, rowData.CODIGO_MATERIAL, rowData.TIPO_ESTRUTURA, rowData.LINHA, rowData.ITEM_COMPONENTE, rowData.QTDE_MONTAGEM, rowData.UNIDADE_MEDIDA, rowData.FATOR_SUCATA];
        ws_data.push(dataRow);
    });
    if (ws_data.length <= 1) { Swal.fire("ℹ️ Nada para Exportar", "A tabela está vazia ou contém apenas linhas sem dados preenchidos.", "info"); return; }
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lista Tecnica");
    const now = new Date();
    const dateStr = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, '0') + "-" + String(now.getDate()).padStart(2, '0') + "_" + String(now.getHours()).padStart(2, '0') + "-" + String(now.getMinutes()).padStart(2, '0') + "-" + String(now.getSeconds()).padStart(2, '0');
    XLSX.writeFile(wb, `Lista_Tecnica_${dateStr}.xlsx`);
    Swal.fire("✅ Exportado!", `A lista foi exportada para 'Lista_Tecnica_${dateStr}.xlsx'.`, "success");
}

function carregarExcel(inputElement) {
    const file = inputElement.files[0];
    if (!file) { Swal.fire("⚠️ Nenhum arquivo selecionado", "Por favor, selecione um arquivo Excel.", "warning"); return; }
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (json.length === 0 || !json[0]) { Swal.fire("⚠️ Arquivo Vazio ou Inválido", "O arquivo Excel está vazio, não contém cabeçalhos ou dados.", "warning"); return; }
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
            UNIDADE_MEDIDA: headers.indexOf("UNIDADE_DE_MEDIDA") !== -1 ? headers.indexOf("UNIDADE_DE_MEDIDA") : headers.indexOf("UNIDADE_MEDIDA"),
            FATOR_SUCATA: headers.indexOf("FATOR_SUCATA")
        };
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

function acaoImportouOuAdicionouLinhas() {
    atualizarSequencias();
    atualizarColunaLinha();
    verificarDuplicatas();
}

async function handlePasteMultipleLines(event) {
    const pastedText = (event.clipboardData || window.clipboardData).getData('text');
    if (!pastedText) return;
    const lines = pastedText.split(/\r?\n/).map(line => line.trim()).filter(line => line !== '');
    if (lines.length === 0) return;
    const possibleHeaders = new Set(['seq', 'codigo_material', 'codigo', 'qtd', 'qtde', 'un', 'unidade_de_medida', 'unidade_medida', 'unidade', 'descricao', 'item_componente', 'item', 'linha', 'nivel', 'site', 'alternativa', 'tipo_estrutura', 'fator_sucata']);
    const normalizeHeader = str => str.normalize('NFD').replace(/[\u0300-\u036f\s_]/g, "").toLowerCase();
    const firstLineNormalized = normalizeHeader(lines[0]);
    const isHeader = Array.from(possibleHeaders).some(header => firstLineNormalized.startsWith(normalizeHeader(header)));
    const realLines = lines.slice(isHeader ? 1 : 0);
    if (realLines.length === 0) return;
    const activeElement = document.activeElement;
    if (activeElement && (activeElement.tagName === 'INPUT' || activeElement.tagName === 'SELECT') && activeElement.closest('#listaTabela')) {
        event.preventDefault();
        const targetRow = activeElement.closest('tr');
        const targetTd = activeElement.closest('td');
        const rowIndex = Array.from(tabela.rows).indexOf(targetRow);
        const columnIndex = Array.from(targetRow.children).indexOf(targetTd);
        let itemsPastedCount = 0;
        for (let i = 0; i < realLines.length; i++) {
            let rowToProcess = tabela.rows[rowIndex + i];
            if (!rowToProcess) { rowToProcess = criarLinhaVazia(); tabela.appendChild(rowToProcess); }
            const inputToUpdate = rowToProcess.children[columnIndex]?.querySelector('input, select');
            if (!inputToUpdate) continue;
            let valueToPaste = realLines[i];
            if (targetTd.classList.contains('codigo-material-col') || targetTd.classList.contains('item-componente-col')) { inputToUpdate.value = valueToPaste.toUpperCase(); } else if (targetTd.classList.contains('qtde-montagem-col')) { inputToUpdate.value = valueToPaste.replace(',', '.'); } else if (targetTd.classList.contains('unidade-medida-col')) { inputToUpdate.value = valueToPaste.toLowerCase(); } else { inputToUpdate.value = valueToPaste.toUpperCase(); }
            inputToUpdate.dispatchEvent(new Event('input', { bubbles: true }));
            const updatedRowData = getLinhaData(rowToProcess);
            if (updatedRowData.CODIGO_MATERIAL && updatedRowData.ITEM_COMPONENTE) {
                const existingDuplicateRow = encontrarLinhaDuplicada(updatedRowData.CODIGO_MATERIAL, updatedRowData.ITEM_COMPONENTE, rowToProcess);
                if (existingDuplicateRow) {
                    const qtdeNova = parseFloat(updatedRowData.QTDE_MONTAGEM) || 0;
                    const existingQtdeInput = existingDuplicateRow.querySelector(".qtde-montagem-col input");
                    if (existingQtdeInput) {
                        const currentQtde = parseFloat(existingQtdeInput.value.replace(',', '.')) || 0;
                        existingQtdeInput.value = (currentQtde + qtdeNova).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
                        existingQtdeInput.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                    if (rowToProcess.parentNode) rowToProcess.remove();
                    itemsPastedCount++;
                } else { itemsPastedCount++; }
            } else { itemsPastedCount++; }
        }
        acaoImportouOuAdicionouLinhas();
    }
}

// ===================================================================================
// --- LISTENERS DE EVENTOS PRINCIPAIS ---
// ===================================================================================

function adicionarListenersDeEventos() {
    document.getElementById("criarListaBtn").addEventListener("click", () => {
        tabela.innerHTML = "";
        criar10Linhas();
        acaoImportouOuAdicionouLinhas();
        Swal.fire("✅ Lista Criada!", "10 novas linhas foram adicionadas.", "success");
    });
    document.getElementById("continuarListaBtn").addEventListener("click", () => {
        criar10Linhas();
        acaoImportouOuAdicionouLinhas();
        Swal.fire("➕ Adicionado", "10 novas linhas foram inseridas.", "success");
    });
    document.getElementById("salvarListaBtn").addEventListener("click", exportarParaExcel);
    document.getElementById("copiarSelecionadoBtn").addEventListener("click", () => {
        const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
        if (linhasSelecionadas.length === 0) { Swal.fire("⚠️ Nada para Copiar", "Nenhuma linha selecionada para cópia.", "warning"); return; }
        cacheCopiado = linhasSelecionadas.map(row => getLinhaData(row));
        Swal.fire("✅ Copiado!", `${cacheCopiado.length} linhas copiadas.`, "success");
    });
    document.getElementById("colarBtn").addEventListener("click", () => {
        if (cacheCopiado.length === 0) { Swal.fire("ℹ️ Nada para Colar", "Nenhum dado copiado.", "info"); return; }
        const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
        const startIndex = linhasSelecionadas.length > 0 ? Array.from(tabela.rows).indexOf(linhasSelecionadas[0]) : tabela.rows.length;
        cacheCopiado.forEach((rowData, i) => {
            const targetRow = tabela.rows[startIndex + i];
            if (targetRow) { preencherLinha(targetRow, rowData); } else { const newRow = criarLinha(rowData); tabela.appendChild(newRow); }
        });
        acaoImportouOuAdicionouLinhas();
        Swal.fire("✅ Colado!", `${cacheCopiado.length} linhas coladas.`, "success");
    });
    document.getElementById("deletarSelecionadosBtn").addEventListener("click", () => {
        const linhasParaDeletar = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
        if (linhasParaDeletar.length === 0) { Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione as linhas para deletar.", "warning"); return; }
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
                linhasParaDeletar.forEach(row => row.remove());
                acaoImportouOuAdicionouLinhas();
                Swal.fire('Deletado!', `${linhasParaDeletar.length} linha(s) foram deletadas.`, 'success');
            }
        });
    });
    document.getElementById("inserirAcimaBtn").addEventListener("click", () => {
        const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
        if (linhasSelecionadas.length === 0) { Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione a(s) linha(s) acima da qual deseja inserir.", "warning"); return; }
        const primeiraLinhaSelecionada = linhasSelecionadas[0];
        const novaLinha = criarLinhaVazia();
        tabela.insertBefore(novaLinha, primeiraLinhaSelecionada);
        acaoImportouOuAdicionouLinhas();
        Swal.fire("⬆️ Inserido", "Uma nova linha foi inserida acima da seleção.", "success");
    });
    document.getElementById("inserirAbaixoBtn").addEventListener("click", () => {
        const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
        if (linhasSelecionadas.length === 0) { Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione a(s) linha(s) abaixo da qual deseja inserir.", "warning"); return; }
        const ultimaLinhaSelecionada = linhasSelecionadas[linhasSelecionadas.length - 1];
        const novaLinha = criarLinhaVazia();
        if (ultimaLinhaSelecionada.nextElementSibling) { tabela.insertBefore(novaLinha, ultimaLinhaSelecionada.nextElementSibling); } else { tabela.appendChild(novaLinha); }
        acaoImportouOuAdicionouLinhas();
        Swal.fire("⬇️ Inserido", "Uma nova linha foi inserida abaixo da seleção.", "success");
    });
    document.getElementById("toggleSeqBtn").addEventListener("click", () => {
        seqAtivo = !seqAtivo;
        document.getElementById("listaTabela").classList.toggle("seq-col-hidden", !seqAtivo);
        document.getElementById("toggleSeqBtn").innerHTML = seqAtivo ? '<span class="material-symbols-outlined">visibility</span> SEQ' : '<span class="material-symbols-outlined">visibility_off</span> SEQ';
    });
    document.getElementById("toggleNivelColBtn").addEventListener("click", () => {
        nivelColVisivel = !nivelColVisivel;
        document.getElementById("listaTabela").classList.toggle("nivel-col-hidden", !nivelColVisivel);
        document.getElementById("toggleNivelColBtn").innerHTML = nivelColVisivel ? '<span class="material-symbols-outlined">layers</span> NÍVEL' : '<span class="material-symbols-outlined">layers_clear</span> NÍVEL';
    });
    document.getElementById("toggleHoverEffectBtn").addEventListener("click", () => {
        hoverEffectAtivo = !hoverEffectAtivo;
        const tableElement = document.getElementById("listaTabela");
        tableElement.classList.toggle("hover-effect", hoverEffectAtivo);
        tableElement.classList.toggle("no-hover-effect", !hoverEffectAtivo);
        document.getElementById("toggleHoverEffectBtn").innerHTML = hoverEffectAtivo ? '<span class="material-symbols-outlined">straighten</span> Régua' : '<span class="material-symbols-outlined">format_line_spacing</span> Régua';
    });
    document.getElementById("clearPaintBtn").addEventListener("click", () => {
        corSelecionada = "";
        demarcarLinha = false;
        removerDemarcacao = false;
        document.getElementById("demarcarLinhaCheckbox").checked = false;
        document.getElementById("removerDemarcacaoCheckbox").checked = false;
        Swal.fire("🎨 Limpeza", "Seleção de cor e modos de demarcação limpos.", "info");
    });
    document.getElementById("demarcarLinhaCheckbox").addEventListener("change", (e) => {
        demarcarLinha = e.target.checked;
        if (demarcarLinha) removerDemarcacao = false;
        document.getElementById("removerDemarcacaoCheckbox").checked = false;
    });
    document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", (e) => {
        removerDemarcacao = e.target.checked;
        if (removerDemarcacao) demarcarLinha = false;
        document.getElementById("demarcarLinhaCheckbox").checked = false;
    });
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
    document.querySelectorAll("#attentionColorButtons .paint-btn").forEach(button => {
        button.addEventListener("click", (e) => {
            corSelecionada = e.target.dataset.color;
            Swal.fire(`🎨 Cor Selecionada`, `Cor ${e.target.textContent.trim()} selecionada.`, "info");
        });
    });
    document.getElementById("ignorarDuplicatasCheckbox").addEventListener("change", (e) => { ignorarDuplicatas = e.target.checked; verificarDuplicatas(); });
    document.getElementById("toggleAllCheckboxesHeader").addEventListener("click", (e) => {
        const isChecked = e.target.checked;
        document.getElementById("toggleAllCheckboxes").checked = isChecked;
        tabela.querySelectorAll(".linha-selecao").forEach(checkbox => { checkbox.checked = isChecked; });
    });
    document.getElementById("toggleAllCheckboxes").addEventListener("click", (e) => {
        const isChecked = e.target.checked;
        document.getElementById("toggleAllCheckboxesHeader").checked = isChecked;
        tabela.querySelectorAll(".linha-selecao").forEach(checkbox => { checkbox.checked = isChecked; });
    });
    document.getElementById("inputFile").addEventListener("change", function (e) {
        carregarExcel(e.target);
        e.target.value = '';
    });
    document.getElementById("findDuplicatesBtn").addEventListener("click", () => {
        resetDuplicateSearchState();
        const currentDuplicates = Array.from(tabela.querySelectorAll('tr.highlight-duplicate'));
        if (currentDuplicates.length === 0) { Swal.fire("ℹ️ Sem Duplicatas", "Não há itens duplicados para buscar na lista.", "info"); return; }
        foundDuplicates = currentDuplicates.sort((a, b) => { return Array.from(tabela.rows).indexOf(a) - Array.from(tabela.rows).indexOf(b); });
        currentDuplicateIndex = 0;
        focusOnDuplicate(foundDuplicates[currentDuplicateIndex]);
    });
    document.addEventListener('keydown', (e) => {
        const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
        if (foundDuplicates.length > 0 && tbodyElement.classList.contains("table-faded")) {
            if (e.key === 'Enter') {
                e.preventDefault();
                currentDuplicateIndex++;
                if (currentDuplicateIndex >= foundDuplicates.length) { currentDuplicateIndex = 0; }
                focusOnDuplicate(foundDuplicates[currentDuplicateIndex]);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                resetDuplicateSearchState();
                Swal.fire("✅ Busca Encerrada", "O modo de busca de duplicatas foi desativado.", "info");
            }
        }
    });
    document.addEventListener('paste', handlePasteMultipleLines);
}

// ===================================================================================
// --- SALVAMENTO AUTOMÁTICO E RESTAURAÇÃO DE SESSÃO ---
// ===================================================================================

function atualizarHorarioBackupDisplay(timestamp) {
    const horarioElement = document.getElementById('ultimo-backup-horario');
    if (horarioElement) {
        if (timestamp) {
            const horarioFormatado = new Date(timestamp).toLocaleTimeString('pt-BR');
            horarioElement.textContent = horarioFormatado;
        } else {
            horarioElement.textContent = '--:--:--';
        }
    }
}

function salvarEstadoLocalmente() {
    const linhas = Array.from(tabela.rows);
    if (linhas.length === 0) {
        localStorage.removeItem('listaTecnicaAutoSave');
        atualizarHorarioBackupDisplay(null);
        return;
    }
    const dadosTabela = linhas.map(row => getLinhaData(row));
    const dadosComTimestamp = {
        timestamp: new Date(),
        data: dadosTabela
    };
    localStorage.setItem('listaTecnicaAutoSave', JSON.stringify(dadosComTimestamp));
    atualizarHorarioBackupDisplay(dadosComTimestamp.timestamp);
}

async function gerenciarInicializacao() {
    const dadosSalvosJSON = localStorage.getItem('listaTecnicaAutoSave');
    let dadosSalvos = null;
    let ultimoBackupTimestamp = null;
    if (dadosSalvosJSON) {
        const objetoSalvo = JSON.parse(dadosSalvosJSON);
        if (objetoSalvo.data && objetoSalvo.timestamp) {
            dadosSalvos = objetoSalvo.data;
            ultimoBackupTimestamp = objetoSalvo.timestamp;
        } else {
            dadosSalvos = objetoSalvo;
        }
    }
    atualizarHorarioBackupDisplay(ultimoBackupTimestamp);
    try {
        if (dadosSalvos) {
            const result = await Swal.fire({
                title: 'Como deseja continuar?',
                text: 'Encontramos um trabalho salvo automaticamente.',
                icon: 'question',
                showConfirmButton: true,
                confirmButtonText: '<i class="material-symbols-outlined" style="vertical-align: sub;">history</i> Restaurar Sessão',
                confirmButtonColor: '#3085d6',
                showDenyButton: true,
                denyButtonText: '<i class="material-symbols-outlined" style="vertical-align: sub;">upload_file</i> Importar Arquivo',
                denyButtonColor: '#5cb85c',
                showCancelButton: true,
                cancelButtonText: '<i class="material-symbols-outlined" style="vertical-align: sub;">add_circle</i> Criar Nova Lista',
                cancelButtonColor: '#d33',
                allowOutsideClick: false,
                allowEscapeKey: false
            });
            if (result.isConfirmed) {
                tabela.innerHTML = "";
                dadosSalvos.forEach(rowData => tabela.appendChild(criarLinha(rowData)));
                acaoImportouOuAdicionouLinhas();
                Swal.fire('Restaurado!', 'Sua sessão anterior foi carregada.', 'success');
            } else if (result.isDenied) {
                document.getElementById('inputFile').click();
                if (tabela.rows.length === 0) {
                    criar10Linhas();
                    acaoImportouOuAdicionouLinhas();
                }
            } else if (result.dismiss === Swal.DismissReason.cancel) {
                localStorage.removeItem('listaTecnicaAutoSave');
                atualizarHorarioBackupDisplay(null);
                tabela.innerHTML = "";
                criar10Linhas();
                acaoImportouOuAdicionouLinhas();
                Swal.fire('Tudo pronto!', 'Uma nova lista foi criada.', 'info');
            }
        } else {
            const result = await Swal.fire({
                title: 'Bem-vindo!',
                text: 'Como deseja começar?',
                icon: 'info',
                showConfirmButton: true,
                confirmButtonText: '<i class="material-symbols-outlined" style="vertical-align: sub;">upload_file</i> Importar Arquivo',
                confirmButtonColor: '#5cb85c',
                showCancelButton: true,
                cancelButtonText: '<i class="material-symbols-outlined" style="vertical-align: sub;">add_circle</i> Criar Nova Lista',
                cancelButtonColor: '#3085d6',
                allowOutsideClick: false,
                allowEscapeKey: false
            });
            if (result.isConfirmed) {
                document.getElementById('inputFile').click();
                if (tabela.rows.length === 0) {
                    criar10Linhas();
                    acaoImportouOuAdicionouLinhas();
                }
            } else if (result.dismiss === Swal.DismissReason.cancel) {
                tabela.innerHTML = "";
                criar10Linhas();
                acaoImportouOuAdicionouLinhas();
            }
        }
    } catch (error) {
        console.error("Erro durante a inicialização:", error);
        tabela.innerHTML = "";
        criar10Linhas();
        acaoImportouOuAdicionouLinhas();
        Swal.fire('Ocorreu um erro', 'Iniciando com uma lista vazia.', 'error');
    }
    configurarInterfaceETimers();
    adicionarListenersDeEventos();
}

function configurarInterfaceETimers() {
    const tabelaElement = document.getElementById("listaTabela");
    if (!seqAtivo) tabelaElement.classList.add("seq-col-hidden");
    if (!nivelColVisivel) tabelaElement.classList.add("nivel-col-hidden");
    if (hoverEffectAtivo) tabelaElement.classList.add("hover-effect");
    else tabelaElement.classList.add("no-hover-effect");
    document.getElementById("toggleSeqBtn").innerHTML = '<span class="material-symbols-outlined">visibility</span> SEQ';
    document.getElementById("toggleNivelColBtn").innerHTML = '<span class="material-symbols-outlined">layers</span> NÍVEL';
    document.getElementById("toggleHoverEffectBtn").innerHTML = '<span class="material-symbols-outlined">straighten</span> Régua';
    setInterval(salvarEstadoLocalmente, 900000);
    console.log("Sistema de salvamento automático ativado (a cada 15 minutos).");
}

// ===================================================================================
// --- FUNCIONALIDADE DE TEMA ESCURO (DRACULA) ---
// ===================================================================================
function setupThemeToggle() {
    const themeToggle = document.getElementById('theme-toggle');
    if (!themeToggle) return;
    const body = document.body;
    const applyTheme = (theme) => {
        if (theme === 'dark') {
            body.classList.add('dark-mode');
            themeToggle.textContent = '☀️';
        } else {
            body.classList.remove('dark-mode');
            themeToggle.textContent = '🌙';
        }
    };
    const savedTheme = localStorage.getItem('theme') || 'light';
    applyTheme(savedTheme);
    themeToggle.addEventListener('click', () => {
        let newTheme;
        if (body.classList.contains('dark-mode')) {
            newTheme = 'light';
        } else {
            newTheme = 'dark';
        }
        applyTheme(newTheme);
        localStorage.setItem('theme', newTheme);
    });
}

// ===================================================================================
// --- INICIALIZAÇÃO DA APLICAÇÃO ---
// ===================================================================================
document.addEventListener("DOMContentLoaded", () => {
    gerenciarInicializacao();
    setupThemeToggle();
});