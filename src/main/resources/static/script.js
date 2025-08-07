/**
 * @file script.js
 * @description Script principal e final para a aplicação de Lista Técnica IFS.
 * @version 12.2 - Verificação e correção final do fluxo de salvamento e restauração de sessão.
 */

// ===================================================================================
// --- VARIÁVEIS GLOBAIS ---
// ===================================================================================
let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true;
let nivelColVisivel = true;
let corSelecionadaInfo = { color: "", level: null };
let demarcarLinha = false;
let removerDemarcacao = false;
let ignorarDuplicatas = false;
let hoverEffectAtivo = true;
const nivelColors = ["#4664cf", "#CD5C5C", "#B3E6B3", "#FFD700", "#8A2BE2", "#FF8C00", "#00CED1", "#FF69B4", "#9ACD32", "#DA70D6"];
const tiposEstrutura = ["Manufatura", "Comprado", ""];
const fatorSucata = ["0", "15", ""];
const alternativas = ["*", ""];
const siteValores = ["1", ""];
let foundDuplicates = [];
let currentDuplicateIndex = -1;

// ===================================================================================
// --- FUNCIONALIDADES DE UI ADICIONAIS ---
// ===================================================================================
function iniciarRelogio() {
    const clockElement = document.getElementById('clock');
    if (!clockElement) return;
    const atualizarHorario = () => clockElement.textContent = new Date().toLocaleTimeString('pt-BR');
    setInterval(atualizarHorario, 1000);
    atualizarHorario();
}

function atualizarContagemDeNiveis() {
    const counts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0 };
    tabela.querySelectorAll("tr").forEach(tr => {
        const nivel = parseInt(tr.querySelector("td.nivel-col input")?.value, 10);
        if (nivel >= 1 && nivel <= 10) counts[nivel]++;
    });
    for (let i = 1; i <= 10; i++) {
        const countSpan = document.getElementById(`count-nivel-${i}`);
        if (countSpan) countSpan.textContent = counts[i];
    }
}

// ===================================================================================
// --- BUSCA DE MATERIAIS ---
// ===================================================================================
function filtrarEMostrarResultados(termo) {
    const resultadosDiv = document.getElementById('resultadosBusca');
    resultadosDiv.innerHTML = '';
    if (typeof listaDeMateriais === 'undefined') {
        resultadosDiv.innerHTML = '<div class="resultado-item"><div class="info"><span>Erro: A lista de materiais não foi carregada.</span></div></div>';
        return;
    }
    if (termo.length < 3) {
        resultadosDiv.innerHTML = '<div class="resultado-item"><div class="info"><span>Digite ao menos 3 caracteres...</span></div></div>';
        return;
    }
    const termoBusca = termo.toLowerCase();
    const resultadosFiltrados = listaDeMateriais.filter(item =>
        item.descricao.toLowerCase().includes(termoBusca) || item.codigo.toLowerCase().includes(termoBusca)
    ).slice(0, 100);
    if (resultadosFiltrados.length === 0) {
        resultadosDiv.innerHTML = '<div class="resultado-item"><div class="info"><span>Nenhum item encontrado.</span></div></div>';
        return;
    }
    resultadosFiltrados.forEach(item => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'resultado-item';
        itemDiv.innerHTML = `<div class="info"><strong>${item.descricao}</strong><span>Código: ${item.codigo}</span></div><button class="copiar-item-btn">Copiar</button>`;
        itemDiv.querySelector('.copiar-item-btn').addEventListener('click', (e) => {
            const button = e.target;
            navigator.clipboard.writeText(item.codigo).then(() => {
                button.textContent = 'Copiado!';
                button.classList.add('copiado');
                setTimeout(() => {
                    button.textContent = 'Copiar';
                    button.classList.remove('copiado');
                }, 2000);
            });
        });
        resultadosDiv.appendChild(itemDiv);
    });
}

// ===================================================================================
// --- FUNCIONALIDADES DA TABELA ---
// ===================================================================================
function acaoImportouOuAdicionouLinhas() {
    atualizarSequencias();
    atualizarColunaLinha();
    verificarDuplicatas();
    atualizarContagemDeNiveis();
}

/**
 * @function handlePaste
 * @description Processa dados colados (ex: do Excel) na tabela.
 * @param {ClipboardEvent} e - O evento de colagem.
 */
function handlePaste(e) {
    if (e.target.tagName !== 'INPUT' || e.target.readOnly) {
        return;
    }
    e.preventDefault();

    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    const pastedRows = text.trim().split(/\r?\n/);

    if (pastedRows.length === 0) return;

    const startRow = e.target.closest('tr');
    const startCellIndex = e.target.closest('td').cellIndex;
    let tableRows = Array.from(tabela.rows);
    let currentRowIndex = tableRows.indexOf(startRow);

    pastedRows.forEach((rowText, rowIndex) => {
        let targetRow = tableRows[currentRowIndex + rowIndex];
        if (!targetRow) {
            targetRow = criarLinhaVazia();
            tabela.appendChild(targetRow);
            tableRows.push(targetRow); // Adiciona a nova linha ao array para referência futura
        }

        const pastedCells = rowText.split('\t');
        pastedCells.forEach((cellText, colIndex) => {
            const targetCellIndex = startCellIndex + colIndex;
            const targetCell = targetRow.cells[targetCellIndex];

            if (targetCell) {
                const input = targetCell.querySelector('input:not([readonly]), select');
                if (input) {
                    input.value = cellText.trim();
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }
        });
    });

    acaoImportouOuAdicionouLinhas();
}

function resetDuplicateSearchState() {
    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
    tbodyElement.classList.remove("table-faded");
    tabela.querySelectorAll('tr.highlight-focused-item').forEach(row => {
        row.classList.remove('highlight-focused-item');
        row.classList.remove('temp-highlight-found');
    });
    verificarDuplicatas();
}

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

async function mostrarAlertaDuplicata(newData, existingRow) {
    const result = await Swal.fire({
        title: '⚠️ Duplicata Encontrada!',
        html: `<p>O item <strong>${newData.ITEM_COMPONENTE}</strong> já existe para o material <strong>${newData.CODIGO_MATERIAL}</strong> na linha <strong>${Array.from(tabela.rows).indexOf(existingRow) + 1}</strong>.</p>`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'Ignorar e Inserir',
        cancelButtonText: 'Cancelar Entrada',
        allowOutsideClick: false,
        allowEscapeKey: false,
        reverseButtons: true
    });
    return result.isConfirmed ? 'ignorar' : 'cancelar';
}

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

        if (!ignorarDuplicatas && currentData.CODIGO_MATERIAL && currentData.ITEM_COMPONENTE && (isCodigoMaterialCol || isItemComponenteCol)) {
            const existingDuplicateRow = encontrarLinhaDuplicada(currentData.CODIGO_MATERIAL, currentData.ITEM_COMPONENTE, currentRow);
            if (existingDuplicateRow) {
                const action = await mostrarAlertaDuplicata(currentData, existingDuplicateRow);
                resetDuplicateSearchState();
                if (action === 'cancelar') e.target.value = "";
            }
        }
        verificarDuplicatas();
        if (td.classList.contains('nivel-col')) {
            aplicarIndentacao(currentRow);
            e.target.dispatchEvent(new Event('change', { bubbles: true }));
            atualizarContagemDeNiveis();
        }
        atualizarColunaLinha();
    });

    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            const currentCellIndex = Array.from(e.target.closest('tr').children).indexOf(e.target.closest('td'));
            const nextRow = e.target.closest('tr').nextElementSibling;
            if (nextRow) {
                const nextInput = nextRow.children[currentCellIndex]?.querySelector('input, select');
                if (nextInput) nextInput.focus();
            } else {
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                acaoImportouOuAdicionouLinhas();
                newRow.children[currentCellIndex]?.querySelector('input, select')?.focus();
            }
        }
    });

    td.appendChild(input);
    return td;
}

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

    select.addEventListener('keydown', (e) => {
         if (e.key === 'Enter') {
            e.preventDefault();
            const currentCellIndex = Array.from(e.target.closest('tr').children).indexOf(e.target.closest('td'));
            const nextRow = e.target.closest('tr').nextElementSibling;
            if (nextRow) {
                const nextInput = nextRow.children[currentCellIndex]?.querySelector('input, select');
                if (nextInput) nextInput.focus();
            } else {
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                acaoImportouOuAdicionouLinhas();
                newRow.children[currentCellIndex]?.querySelector('input, select')?.focus();
            }
        }
    });
    td.appendChild(select);
    return td;
}

function criarLinha(v = {}) {
    const row = document.createElement("tr");

    row.addEventListener("click", async (e) => {
        if (removerDemarcacao) {
            const targetRow = e.currentTarget;
            targetRow.style.backgroundColor = "";
            for (let i = 1; i <= 10; i++) targetRow.classList.remove(`nivel-${i}`);
            const nivelInput = targetRow.querySelector(".nivel-col input");
            if (nivelInput) {
                nivelInput.value = "";
                nivelInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
            return;
        }
        if (demarcarLinha && corSelecionadaInfo.color) {
            const targetRow = e.currentTarget;
            targetRow.style.backgroundColor = corSelecionadaInfo.color;
            if (corSelecionadaInfo.level) {
                const nivelInput = targetRow.querySelector(".nivel-col input");
                const nivelAtual = nivelInput.value.trim();
                const novoNivel = String(corSelecionadaInfo.level);
                if (nivelAtual && nivelAtual !== novoNivel) {
                    const result = await Swal.fire({
                        title: 'Alterar Nível?',
                        html: `A linha já possui o Nível <strong>${nivelAtual}</strong>. Deseja alterar para o Nível <strong>${novoNivel}</strong>?`,
                        icon: 'question', showCancelButton: true, confirmButtonText: 'Sim, alterar',
                        cancelButtonText: 'Cancelar', confirmButtonColor: '#28a745', cancelButtonColor: '#d33'
                    });
                    if (!result.isConfirmed) {
                        targetRow.style.backgroundColor = "";
                        return;
                    }
                }
                nivelInput.value = novoNivel;
                nivelInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
        }
    });

    const checkboxTd = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.classList.add("linha-selecao");
    checkboxTd.appendChild(checkbox);
    row.appendChild(checkboxTd);
    const seqTd = document.createElement("td");
    seqTd.classList.add("seq-col");
    row.appendChild(seqTd);
    row.appendChild(inputCell("text", false, v.NIVEL || "", true, "nivel-col"));
    row.appendChild(selectCell(siteValores, v.SITE || "1"));
    row.appendChild(selectCell(alternativas, v.ALTERNATIVA || "*"));
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || "", true, "codigo-material-col"));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));
    row.appendChild(inputCell("text", true, v.LINHA || "", false, "linha-auto-col"));
    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || "", true, "item-componente-col"));
    row.appendChild(inputCell("text", false, v.QTDE_MONTAGEM || "", false, "qtde-montagem-col"));
    row.appendChild(inputCell("text", false, (v.UNIDADE_MEDIDA || "").toLowerCase(), true, "unidade-medida-col"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    aplicarIndentacao(row);

    if (!seqAtivo) seqTd.style.display = "none";

    const nivelCell = row.querySelector('.nivel-col');
    if (nivelCell && !nivelColVisivel) nivelCell.style.display = "none";

    return row;
}

function criarLinhaVazia() { return criarLinha({}); }
function criar10Linhas() { for (let i = 0; i < 10; i++) tabela.appendChild(criarLinhaVazia()); }

function atualizarSequencias() {
    tabela.querySelectorAll("tr").forEach((row, index) => {
        row.querySelectorAll("td")[1].textContent = (index + 1);
    });
}

function atualizarColunaLinha() {
    let currentCodigoMaterial = "";
    let currentSequence = 10;
    tabela.querySelectorAll("tr").forEach(row => {
        const codigoMaterial = row.querySelector(".codigo-material-col input")?.value.trim();
        const linhaInput = row.querySelector(".linha-auto-col input");
        if (linhaInput) {
            if (codigoMaterial) {
                if (codigoMaterial !== currentCodigoMaterial) {
                    currentCodigoMaterial = codigoMaterial;
                    currentSequence = 10;
                }
                linhaInput.value = String(currentSequence);
                currentSequence += 10;
            } else {
                linhaInput.value = "";
                currentCodigoMaterial = "";
            }
        }
    });
}

function aplicarIndentacao(row) {
    for (let i = 1; i <= 10; i++) row.classList.remove(`nivel-${i}`);
    const nivel = parseInt(row.querySelector(".nivel-col input")?.value);
    if (!isNaN(nivel) && nivel >= 1 && nivel <= 10) row.classList.add(`nivel-${nivel}`);
}

function getLinhaData(tr) {
    const cells = tr.querySelectorAll("td");
    return {
        NIVEL: cells[2]?.querySelector("input")?.value.trim() || "", SITE: cells[3]?.querySelector("select")?.value || "",
        ALTERNATIVA: cells[4]?.querySelector("select")?.value || "", CODIGO_MATERIAL: cells[5]?.querySelector("input")?.value.trim().toUpperCase() || "",
        TIPO_ESTRUTURA: cells[6]?.querySelector("select")?.value || "", LINHA: cells[7]?.querySelector("input")?.value || "",
        ITEM_COMPONENTE: cells[8]?.querySelector("input")?.value.trim().toUpperCase() || "", QTDE_MONTAGEM: cells[9]?.querySelector("input")?.value.trim() || "",
        UNIDADE_MEDIDA: cells[10]?.querySelector("input")?.value.trim().toLowerCase() || "", FATOR_SUCATA: cells[11]?.querySelector("select")?.value || ""
    };
}

function preencherLinha(targetRow, rowData) {
    const inputs = targetRow.querySelectorAll("input, select");
    if(inputs.length < 11) return;
    inputs[1].value = rowData.NIVEL || "";
    inputs[2].value = rowData.SITE || "1";
    inputs[3].value = rowData.ALTERNATIVA || "*";
    inputs[4].value = rowData.CODIGO_MATERIAL || "";
    inputs[5].value = rowData.TIPO_ESTRUTURA || "Manufatura";
    inputs[7].value = rowData.ITEM_COMPONENTE || "";
    inputs[8].value = rowData.QTDE_MONTAGEM || "";
    inputs[9].value = rowData.UNIDADE_MEDIDA || "";
    inputs[10].value = rowData.FATOR_SUCATA || "0";
    inputs.forEach(input => input.dispatchEvent(new Event('input', { bubbles: true })));
}

function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    linhas.forEach(row => row.classList.remove("highlight-duplicate"));
    const combinaçõesDetectadas = new Map();
    const tempFoundDuplicates = [];
    if (ignorarDuplicatas) {
        document.getElementById("duplicateCountDisplay").textContent = "";
        foundDuplicates = [];
        return;
    }
    linhas.forEach(tr => {
        const data = getLinhaData(tr);
        if (data.CODIGO_MATERIAL === "" || data.ITEM_COMPONENTE === "") return;
        const hash = `${data.CODIGO_MATERIAL}|${data.ITEM_COMPONENTE}`;
        if (!combinaçõesDetectadas.has(hash)) combinaçõesDetectadas.set(hash, []);
        combinaçõesDetectadas.get(hash).push(tr);
    });
    for (const rows of combinaçõesDetectadas.values()) {
        if (rows.length > 1) {
            rows.forEach(row => {
                row.classList.add("highlight-duplicate");
                tempFoundDuplicates.push(row);
            });
        }
    }
    foundDuplicates = tempFoundDuplicates.sort((a, b) => Array.from(tabela.rows).indexOf(a) - Array.from(tabela.rows).indexOf(b));
    document.getElementById("duplicateCountDisplay").textContent = foundDuplicates.length > 0 ? `⚠️ ${foundDuplicates.length} duplicata(s)` : "";
}

function getContrastColor(hexcolor) {
    if (!hexcolor.startsWith("#")) return "black";
    const r = parseInt(hexcolor.slice(1, 3), 16), g = parseInt(hexcolor.slice(3, 5), 16), b = parseInt(hexcolor.slice(5, 7), 16);
    return (Math.sqrt(0.299 * (r * r) + 0.587 * (g * g) + 0.114 * (b * b)) > 127.5) ? "black" : "white";
}

function exportarParaExcel() {
    const ws_data = [["NIVEL", "SITE", "ALTERNATIVA", "CODIGO_MATERIAL", "TIPO ESTRUTURA", "LINHA", "ITEM_COMPONENTE", "QTDE_MONTAGEM", "UNIDADE DE MEDIDA", "FATOR_SUCATA"]];
    tabela.querySelectorAll("tr").forEach(row => {
        const rowData = getLinhaData(row);
        if (rowData.CODIGO_MATERIAL === "" && rowData.ITEM_COMPONENTE === "") return;
        ws_data.push(Object.values(rowData));
    });
    if (ws_data.length <= 1) { Swal.fire("ℹ️ Nada para Exportar", "A tabela está vazia.", "info"); return; }
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lista Tecnica");
    const now = new Date();
    const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}-${String(now.getSeconds()).padStart(2, '0')}`;
    const fileName = `Lista_Tecnica_${dateStr}.xlsx`;
    XLSX.writeFile(wb, fileName);
    let successMessage = `A lista foi exportada para '${fileName}'.`;
    const bannerText = document.getElementById('editable-banner').textContent.trim();
    const defaultBannerText = 'Clique para editar uma mensagem ou anotação aqui...';
    if (bannerText && bannerText !== defaultBannerText) {
        successMessage += `<br><br><strong>Anotação:</strong><br><em>${bannerText}</em>`;
    }
    Swal.fire({ title: "✅ Exportado!", html: successMessage, icon: "success" });
}

function carregarExcel(inputElement) {
    const file = inputElement.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (json.length < 2) return;
        tabela.innerHTML = "";
        const headers = json[0].map(h => String(h).trim().replace(/\s/g, '_').toUpperCase());
        const colIndices = {
            NIVEL: headers.indexOf("NIVEL"), SITE: headers.indexOf("SITE"),
            ALTERNATIVA: headers.indexOf("ALTERNATIVA"), CODIGO_MATERIAL: headers.indexOf("CODIGO_MATERIAL"),
            TIPO_ESTRUTURA: headers.indexOf("TIPO_ESTRUTURA"), ITEM_COMPONENTE: headers.indexOf("ITEM_COMPONENTE"),
            QTDE_MONTAGEM: headers.indexOf("QTDE_MONTAGEM"), UNIDADE_MEDIDA: headers.indexOf("UNIDADE_DE_MEDIDA") !== -1 ? headers.indexOf("UNIDADE_DE_MEDIDA") : headers.indexOf("UNIDADE_MEDIDA"),
            FATOR_SUCATA: headers.indexOf("FATOR_SUCATA")
        };
        json.slice(1).forEach(rowData => {
            const rowObj = {
                NIVEL: rowData[colIndices.NIVEL] ?? "", SITE: rowData[colIndices.SITE] ?? "1",
                ALTERNATIVA: rowData[colIndices.ALTERNATIVA] ?? "*", CODIGO_MATERIAL: rowData[colIndices.CODIGO_MATERIAL] ?? "",
                TIPO_ESTRUTURA: rowData[colIndices.TIPO_ESTRUTURA] ?? "Manufatura", ITEM_COMPONENTE: rowData[colIndices.ITEM_COMPONENTE] ?? "",
                QTDE_MONTAGEM: rowData[colIndices.QTDE_MONTAGEM] ?? "", UNIDADE_MEDIDA: String(rowData[colIndices.UNIDADE_MEDIDA] ?? "").toLowerCase(),
                FATOR_SUCATA: rowData[colIndices.FATOR_SUCATA] ?? "0"
            };
            tabela.appendChild(criarLinha(rowObj));
        });
        acaoImportouOuAdicionouLinhas();
        Swal.fire("✅ Importado!", "Os dados do Excel foram carregados.", "success");
    };
    reader.readAsArrayBuffer(file);
}

function salvarLocalManualmente() {
    salvarEstadoLocalmente();
    Swal.fire({
        toast: true,
        position: 'top-end',
        icon: 'success',
        title: 'Sessão salva com sucesso!',
        showConfirmButton: false,
        timer: 2000,
        timerProgressBar: true
    });
}

function iniciarBuscaDeDuplicatas() {
    verificarDuplicatas();
    if (foundDuplicates.length === 0) {
        Swal.fire("ℹ️ Sem Duplicatas", "Não há itens duplicados para buscar na lista.", "info");
        return;
    }
    currentDuplicateIndex = 0;
    focusOnDuplicate(foundDuplicates[currentDuplicateIndex]);
}

function focusOnDuplicate(rowToFocus) {
    if (!rowToFocus) return;
    const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
    tabela.querySelectorAll('tr').forEach(row => row.classList.remove('highlight-focused-item', 'temp-highlight-found'));
    tbodyElement.classList.add("table-faded");
    rowToFocus.classList.add('highlight-focused-item', 'highlight-duplicate', 'temp-highlight-found');
    rowToFocus.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function adicionarListenersDeEventos() {
    tabela.addEventListener('paste', handlePaste);

    document.getElementById("criarListaBtn").addEventListener("click", () => {
        if (tabela.rows.length > 0 && Array.from(tabela.rows).some(row => getLinhaData(row).CODIGO_MATERIAL)) {
            Swal.fire({ title: 'Criar Nova Lista?', text: "Qualquer trabalho não salvo será perdido.", icon: 'warning', showCancelButton: true, confirmButtonText: 'Sim, criar nova', cancelButtonText: 'Cancelar' }).then(result => {
                if (result.isConfirmed) {
                    tabela.innerHTML = "";
                    criar10Linhas();
                    acaoImportouOuAdicionouLinhas();
                }
            });
        } else {
            tabela.innerHTML = "";
            criar10Linhas();
            acaoImportouOuAdicionouLinhas();
        }
    });
    document.getElementById("continuarListaBtn").addEventListener("click", () => {
        criar10Linhas();
        acaoImportouOuAdicionouLinhas();
    });
    document.getElementById("salvarListaBtn").addEventListener("click", exportarParaExcel);
    document.getElementById("salvarLocalBtn").addEventListener("click", salvarLocalManualmente);
    document.getElementById("deletarSelecionadosBtn").addEventListener("click", () => {
        const linhasParaDeletar = Array.from(tabela.querySelectorAll(".linha-selecao:checked"));
        if (linhasParaDeletar.length === 0) { Swal.fire("⚠️ Nenhuma Linha Selecionada", "", "warning"); return; }
        Swal.fire({ title: 'Tem certeza?', text: `Você vai deletar ${linhasParaDeletar.length} linha(s).`, icon: 'warning', showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Sim, deletar!' }).then((result) => {
            if (result.isConfirmed) {
                linhasParaDeletar.forEach(cb => cb.closest("tr").remove());
                acaoImportouOuAdicionouLinhas();
            }
        });
    });
    document.getElementById("inserirAcimaBtn").addEventListener("click", () => {
        const primeiraLinhaSelecionada = tabela.querySelector(".linha-selecao:checked")?.closest("tr");
        if (!primeiraLinhaSelecionada) { Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione a linha acima da qual deseja inserir.", "warning"); return; }
        tabela.insertBefore(criarLinhaVazia(), primeiraLinhaSelecionada);
        acaoImportouOuAdicionouLinhas();
    });
    document.getElementById("inserirAbaixoBtn").addEventListener("click", () => {
        const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
        if (linhasSelecionadas.length === 0) { Swal.fire("⚠️ Nenhuma Linha Selecionada", "Selecione a linha abaixo da qual deseja inserir.", "warning"); return; }
        const ultimaLinhaSelecionada = linhasSelecionadas[linhasSelecionadas.length - 1];
        ultimaLinhaSelecionada.insertAdjacentElement('afterend', criarLinhaVazia());
        acaoImportouOuAdicionouLinhas();
    });
    document.getElementById("toggleSeqBtn").addEventListener("click", () => { seqAtivo = !seqAtivo; document.getElementById("listaTabela").classList.toggle("seq-col-hidden", !seqAtivo); });
    document.getElementById("toggleNivelColBtn").addEventListener("click", () => { nivelColVisivel = !nivelColVisivel; document.getElementById("listaTabela").classList.toggle("nivel-col-hidden", !nivelColVisivel); });
    document.getElementById("toggleHoverEffectBtn").addEventListener("click", () => { hoverEffectAtivo = !hoverEffectAtivo; document.getElementById("listaTabela").classList.toggle("hover-effect", !hoverEffectAtivo); });
    document.getElementById("clearPaintBtn").addEventListener("click", () => { corSelecionadaInfo = { color: "", level: null }; });
    document.getElementById("demarcarLinhaCheckbox").addEventListener("change", (e) => demarcarLinha = e.target.checked);
    document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", (e) => removerDemarcacao = e.target.checked);
    const nivelColorButtonsDiv = document.getElementById("nivelColorButtons");
    nivelColors.forEach((color, index) => {
        const button = document.createElement("button");
        button.className = "paint-btn";
        button.style.backgroundColor = color;
        button.style.color = getContrastColor(color);
        button.textContent = `Nível ${index + 1}`;
        button.addEventListener("click", () => {
            corSelecionadaInfo = { color: color, level: index + 1 };
            Swal.fire({ toast: true, position: 'top-end', icon: 'info', title: `Cor do Nível ${index + 1} selecionada.`, showConfirmButton: false, timer: 1500 });
        });
        nivelColorButtonsDiv.appendChild(button);
    });
    document.querySelectorAll("#attentionColorButtons .paint-btn").forEach(button => {
        button.addEventListener("click", () => {
            corSelecionadaInfo = { color: button.dataset.color, level: null };
            Swal.fire({ toast: true, position: 'top-end', icon: 'info', title: `Cor de ${button.textContent.trim()} selecionada.`, showConfirmButton: false, timer: 1500 });
        });
    });
    document.getElementById("ignorarDuplicatasCheckbox").addEventListener("change", (e) => { ignorarDuplicatas = e.target.checked; verificarDuplicatas(); });
    document.getElementById("toggleAllCheckboxesHeader").addEventListener("change", (e) => {
        document.querySelectorAll("#listaTabela .linha-selecao, #toggleAllCheckboxes").forEach(cb => cb.checked = e.target.checked);
    });
    document.getElementById("toggleAllCheckboxes").addEventListener("change", (e) => {
        document.querySelectorAll("#listaTabela .linha-selecao, #toggleAllCheckboxesHeader").forEach(cb => cb.checked = e.target.checked);
    });
    document.getElementById("inputFile").addEventListener("change", function(e) {
        carregarExcel(e.target);
        e.target.value = '';
    });
    document.getElementById("findDuplicatesBtn").addEventListener("click", () => {
        if (ignorarDuplicatas) {
            Swal.fire({ title: 'Busca Desativada', text: "A opção 'Ignorar duplicatas' está ativa. Deseja desativá-la e buscar?", icon: 'info', showCancelButton: true, confirmButtonText: 'Sim, buscar' }).then(result => {
                if (result.isConfirmed) {
                    document.getElementById('ignorarDuplicatasCheckbox').checked = false;
                    ignorarDuplicatas = false;
                    iniciarBuscaDeDuplicatas();
                }
            });
        } else {
            iniciarBuscaDeDuplicatas();
        }
    });
    const modal = document.getElementById('modalBusca'), buscarItemBtn = document.getElementById('buscarItemBtn'), fecharModalBtn = document.getElementById('fecharModalBtn'), inputBusca = document.getElementById('inputBuscaItem');
    buscarItemBtn.addEventListener('click', () => { modal.style.display = 'flex'; inputBusca.focus(); });
    fecharModalBtn.addEventListener('click', () => { modal.style.display = 'none'; });
    modal.addEventListener('click', (e) => { if (e.target === modal) modal.style.display = 'none'; });
    inputBusca.addEventListener('input', () => filtrarEMostrarResultados(inputBusca.value));

    document.addEventListener('keydown', (e) => {
        const tbodyElement = document.getElementById("listaTabela").querySelector('tbody');
        if (foundDuplicates.length > 0 && tbodyElement.classList.contains("table-faded")) {
            if (e.key === 'Enter') {
                e.preventDefault();
                currentDuplicateIndex = (currentDuplicateIndex + 1) % foundDuplicates.length;
                focusOnDuplicate(foundDuplicates[currentDuplicateIndex]);
                return;
            }
            if (e.key === 'Escape') {
                e.preventDefault();
                resetDuplicateSearchState();
                return;
            }
        }

        if (e.ctrlKey && e.key === '1') {
            e.preventDefault();
            document.getElementById('salvarLocalBtn').click();
            return;
        }

        if (e.ctrlKey && e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
            e.preventDefault();
            document.getElementById(e.key === 'ArrowUp' ? 'inserirAcimaBtn' : 'inserirAbaixoBtn').click();
            return;
        }

        // Trata Ctrl + Delete como atalho para deletarSelecionadosBtn
        if (e.ctrlKey && e.key === 'Delete') {
            e.preventDefault();
            document.getElementById('deletarSelecionadosBtn')?.click();
            return;
        }

        if (e.target.tagName === 'INPUT' || e.target.isContentEditable) return;

        const shortcuts = {
            '\'': 'salvarListaBtn',
            'g': 'continuarListaBtn',
            'n': 'criarListaBtn',
            'm': 'buscarItemBtn'
            // 'd' REMOVIDO DAQUI
        };

        if (e.ctrlKey && shortcuts[e.key.toLowerCase()]) {
            e.preventDefault();
            document.getElementById(shortcuts[e.key.toLowerCase()]).click();
        }

        if (e.ctrlKey && e.key.toLowerCase() === 'i') {
            e.preventDefault();
            document.querySelector('label[for="inputFile"]').click();
        }

        if (e.ctrlKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
            e.preventDefault();
            const row = tabela.querySelector(e.key === 'ArrowUp' ? 'tr:first-child' : 'tr:last-child');
            row?.querySelector('input, select')?.focus();
        }
    });

}

// ===================================================================================
// --- INICIALIZAÇÃO E SESSÃO ---
// ===================================================================================
function atualizarHorarioBackupDisplay(timestamp) {
    document.getElementById('ultimo-backup-horario').textContent = timestamp ? new Date(timestamp).toLocaleTimeString('pt-BR') : '--:--:--';
}
function salvarEstadoLocalmente() {
    try {
        const dadosTabela = Array.from(tabela.rows).map(getLinhaData);
        if (dadosTabela.some(d => d.CODIGO_MATERIAL || d.ITEM_COMPONENTE)) {
            const dadosComTimestamp = { timestamp: new Date().toISOString(), data: dadosTabela };
            localStorage.setItem('listaTecnicaAutoSave', JSON.stringify(dadosComTimestamp));
            atualizarHorarioBackupDisplay(dadosComTimestamp.timestamp);
        } else {
            localStorage.removeItem('listaTecnicaAutoSave');
            atualizarHorarioBackupDisplay(null);
        }
    } catch (error) {
        console.error("Erro ao salvar estado localmente:", error);
    }
}

async function gerenciarInicializacao() {
    const dadosSalvosJSON = localStorage.getItem('listaTecnicaAutoSave');
    let dadosSalvos = null, ultimoBackupTimestamp = null;
    if (dadosSalvosJSON) {
        try {
            const objetoSalvo = JSON.parse(dadosSalvosJSON);
            if (objetoSalvo && objetoSalvo.data) {
                dadosSalvos = objetoSalvo.data;
                ultimoBackupTimestamp = objetoSalvo.timestamp;
            }
        } catch (error) {
            console.error("Erro ao parsear dados salvos. Removendo item corrompido.", error);
            localStorage.removeItem('listaTecnicaAutoSave');
        }
    }
    atualizarHorarioBackupDisplay(ultimoBackupTimestamp);

    if (dadosSalvos && dadosSalvos.length > 0) {
        const result = await Swal.fire({
            title: 'Como deseja continuar?',
            html: `Encontramos uma sessão salva de <strong>${new Date(ultimoBackupTimestamp).toLocaleString('pt-BR')}</strong>.`,
            icon: 'question',
            showConfirmButton: true, confirmButtonText: 'Restaurar Sessão', confirmButtonColor: '#3085d6',
            showDenyButton: true, denyButtonText: 'Importar Arquivo', denyButtonColor: '#28a745',
            showCancelButton: true, cancelButtonText: 'Criar Nova Lista', cancelButtonColor: '#a31f35',
            allowOutsideClick: false
        });
        if (result.isConfirmed) {
            tabela.innerHTML = "";
            dadosSalvos.forEach(rowData => tabela.appendChild(criarLinha(rowData)));
            acaoImportouOuAdicionouLinhas();
        } else if (result.isDenied) {
            document.getElementById('inputFile').click();
        } else if (result.dismiss === Swal.DismissReason.cancel) {
            localStorage.removeItem('listaTecnicaAutoSave');
            atualizarHorarioBackupDisplay(null);
            tabela.innerHTML = "";
            criar10Linhas();
            acaoImportouOuAdicionouLinhas();
        }
    } else {
        const result = await Swal.fire({
            title: 'Bem-vindo!', text: 'Como deseja começar?', icon: 'info',
            showConfirmButton: true, confirmButtonText: 'Importar Lista', confirmButtonColor: '#28a745',
            showCancelButton: true, cancelButtonText: 'Criar Nova Lista', cancelButtonColor: '#a31f35',
            allowOutsideClick: false
        });
        if (result.isConfirmed) {
            document.getElementById('inputFile').click();
        } else if (result.dismiss === Swal.DismissReason.cancel) {
            tabela.innerHTML = "";
            criar10Linhas();
            acaoImportouOuAdicionouLinhas();
        }
    }
    adicionarListenersDeEventos();
    configurarInterfaceETimers();
}

function configurarInterfaceETimers() {
    setInterval(salvarEstadoLocalmente, 60000); // Salva a cada 1 minuto
}

function setupThemeToggle() {
    const themeToggle = document.getElementById('theme-toggle');
    const body = document.body;
    const applyTheme = (theme) => {
        body.classList.toggle('dark-mode', theme === 'dark');
        themeToggle.textContent = theme === 'dark' ? '☀️' : '🌙';
    };
    const savedTheme = localStorage.getItem('theme') || 'light';
    applyTheme(savedTheme);
    themeToggle.addEventListener('click', () => {
        const newTheme = body.classList.contains('dark-mode') ? 'light' : 'dark';
        applyTheme(newTheme);
        localStorage.setItem('theme', newTheme);
    });
}

document.addEventListener("DOMContentLoaded", () => {
    iniciarRelogio();
    gerenciarInicializacao();
    setupThemeToggle();
    atualizarContagemDeNiveis();
});