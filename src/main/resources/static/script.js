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
    input.addEventListener("input", (e) => {
        // Converte para minúsculas se for a coluna de unidade de medida, maiúsculas caso contrário
        if (e.target.closest('td').classList.contains('unidade-medida-col')) {
            e.target.value = e.target.value.toLowerCase();
        } else {
            e.target.value = e.target.value.toUpperCase();
        }
        verificarDuplicatas(); // Verifica duplicatas a cada alteração
        if (td.classList.contains('nivel-col')) {
            aplicarIndentacao(e.target.closest('tr')); // Atualiza indentação se o NÍVEL muda
            e.target.dispatchEvent(new Event('change', { bubbles: true })); // Dispara evento change para outras lógicas
        }
        atualizarColunaLinha(); // Sempre atualiza LINHA se campos relevantes mudam
    });

    // Navegação com Enter
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault(); // Previne quebra de linha padrão
            const currentInput = e.target;
            const currentTd = currentInput.closest('td');
            const currentRow = currentInput.closest('tr');
            const currentRowIndex = Array.from(tabela.rows).indexOf(currentRow);
            const currentCellIndex = Array.from(currentRow.children).indexOf(currentTd);

            const nextRow = tabela.rows[currentRowIndex + 1];
            if (nextRow) {
                // Tenta focar na mesma coluna na próxima linha
                const nextTd = nextRow.children[currentCellIndex];
                const nextInput = nextTd?.querySelector('input, select');
                if (nextInput) {
                    nextInput.focus();
                } else {
                    // Se não houver input na mesma coluna na próxima linha, tenta a próxima célula na linha atual
                    const nextCellInRow = currentRow.children[currentCellIndex + 1];
                    const nextInputInRow = nextCellInRow?.querySelector('input, select');
                    if (nextInputInRow) {
                        nextInputInRow.focus();
                    }
                }
            } else {
                // Se for a última linha, cria uma nova linha e foca na mesma coluna dela
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                acaoImportouOuAdicionouLinhas(); // Atualiza SEQ e LINHA após adicionar nova linha
                const firstInputInNewRow = newRow.children[currentCellIndex]?.querySelector('input, select');
                if (firstInputInNewRow) {
                    firstInputInNewRow.focus();
                }
            }
        }
    });

    // Funcionalidade de pintura/demarcação de célula
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

    // Funcionalidade de pintura/demarcação de célula
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

    // Navegação com Enter
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
                acaoImportouOuAdicionouLinhas(); // Atualiza SEQ e LINHA após adicionar nova linha
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
        // Ignora cliques em inputs, selects ou botões dentro da linha
        if (e.target.tagName.match(/INPUT|SELECT|BUTTON/)) return;
        if (demarcarLinha) { // Apenas se o modo "Demarcar linha" estiver ativo
            if (removerDemarcacao) {
                row.style.backgroundColor = ""; // Remove cor
            } else if (corSelecionada) {
                // Alterna a cor: se já tem a cor selecionada, remove; senão, aplica
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
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || "", true));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));

    // Célula para LINHA - agora é um input somente leitura e preenchido automaticamente
    const linhaCell = inputCell("text", true, v.LINHA || "");
    linhaCell.classList.add("linha-auto-col");
    row.appendChild(linhaCell);

    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || "", true));
    row.appendChild(inputCell("text", false, v.QTDE_MONTAGEM || ""));
    // AQUI: Garantir que UNIDADE_MEDIDA use o valor corretamente
    row.appendChild(inputCell("text", false, (v.UNIDADE_MEDIDA || "").toLowerCase(), true, "unidade-medida-col"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    // Aplica a indentação visual baseada no nível
    aplicarIndentacao(row);

    // Esconde colunas se não estiverem ativas
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
        const seqTd = row.querySelectorAll("td")[1]; // Índice 1 é a coluna SEQ
        if (seqTd) seqTd.textContent = (index + 1) * 10;
    });
}

/**
 * Preenche automaticamente a coluna "LINHA" baseando-se no `CODIGO_MATERIAL` e `ITEM_COMPONENTE`.
 * Se um grupo de `CODIGO_MATERIAL` contiver qualquer `ITEM_COMPONENTE` que comece com "MP-",
 * todas as linhas desse grupo terão a "LINHA" definida como "10".
 * Caso contrário, a "LINHA" será uma sequência incremental (10, 20, 30...).
 */
function atualizarColunaLinha() {
    const rows = Array.from(tabela.rows);
    const groupedData = new Map(); // Mapa para armazenar grupos de linhas por CODIGO_MATERIAL

    // Primeira passagem: Agrupar linhas e detectar "MP-"
    rows.forEach(row => {
        const data = getLinhaData(row);
        const codigoMaterial = data.CODIGO_MATERIAL.trim();
        const itemComponente = data.ITEM_COMPONENTE.trim();

        // Se a linha não tem CODIGO_MATERIAL nem ITEM_COMPONENTE, ela é "vazia" para o cálculo da LINHA
        if (codigoMaterial === "" && itemComponente === "") {
            const linhaInput = row.querySelectorAll("td")[7]?.querySelector("input");
            if (linhaInput) linhaInput.value = ""; // Garante que a LINHA fica vazia
            return;
        }

        // Se tem CODIGO_MATERIAL, agrupa por ele
        if (codigoMaterial !== "") {
            if (!groupedData.has(codigoMaterial)) {
                groupedData.set(codigoMaterial, {
                    rows: [],
                    hasMP: false
                });
            }
            const group = groupedData.get(codigoMaterial);
            group.rows.push(row);
            if (itemComponente.toUpperCase().startsWith("MP-")) {
                group.hasMP = true;
            }
        }
        // Caso a linha tenha ITEM_COMPONENTE mas não CODIGO_MATERIAL, ela ainda não pertence a um grupo
        // e terá sua LINHA vazia, a menos que uma lógica específica para isso seja implementada.
        // Pelo requisito, parece que CODIGO_MATERIAL é o principal agrupador.
        // Se ela não for agrupada, será tratada no loop abaixo.
    });

    // Segunda passagem: Atribuir valores de LINHA
    let currentCodigoMaterial = "";
    let currentSequence = 10;

    rows.forEach(row => {
        const data = getLinhaData(row);
        const codigoMaterial = data.CODIGO_MATERIAL.trim();
        const itemComponente = data.ITEM_COMPONENTE.trim();
        const linhaInput = row.querySelectorAll("td")[7]?.querySelector("input");

        if (!linhaInput) return; // Se não encontrou o input da LINHA, pula a linha

        // Se a linha é um "separador" (vazia), sua LINHA já foi definida como vazia e reiniciamos a sequência
        if (codigoMaterial === "" && itemComponente === "") {
            currentCodigoMaterial = ""; // Reseta o material atual
            currentSequence = 10; // Reinicia a sequência
            // A LINHA já está vazia, não precisa fazer mais nada
            return;
        }

        // Se o CODIGO_MATERIAL mudou, ou é o primeiro item de um novo grupo
        if (codigoMaterial !== "" && codigoMaterial !== currentCodigoMaterial) {
            currentCodigoMaterial = codigoMaterial; // Define o novo grupo
            currentSequence = 10; // Reinicia a sequência para este novo grupo
        }

        // Se a linha pertence a um grupo com CODIGO_MATERIAL
        if (codigoMaterial !== "") {
            const group = groupedData.get(codigoMaterial);
            if (group && group.hasMP) {
                linhaInput.value = "10";
            } else {
                linhaInput.value = String(currentSequence);
                currentSequence += 10;
            }
        } else {
            // Se a linha tem ITEM_COMPONENTE mas não CODIGO_MATERIAL, ou se a lógica não a agrupou
            // (e não é um "separador" totalmente vazio), sua LINHA pode ser vazia por padrão
            // ou seguir uma sequência independente se for um caso de uso futuro.
            // Pelo requisito atual, apenas linhas com CODIGO_MATERIAL definem grupos para a regra "MP-".
            // Para as outras, manter vazio por padrão.
            linhaInput.value = "";
            currentSequence = 10; // Reseta a sequência, pois este item não faz parte de um grupo numérico
            currentCodigoMaterial = "";
        }
    });
}

/**
 * Aplica classes CSS à linha para indentação visual baseada no valor da coluna "NÍVEL".
 * @param {HTMLTableRowElement} row - A linha (<tr>) a ser indentada.
 */
function aplicarIndentacao(row) {
    // Remove todas as classes de nível existentes
    for (let i = 1; i <= 10; i++) row.classList.remove(`nivel-${i}`);
    const nivelInput = row.querySelectorAll("td")[2]?.querySelector("input"); // Índice 2 é a coluna NÍVEL
    if (nivelInput) {
        let nivel = parseInt(nivelInput.value);
        if (!isNaN(nivel) && nivel >= 1 && nivel <= 10) {
            row.classList.add(`nivel-${nivel}`); // Adiciona a classe correspondente ao nível
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
        LINHA: cells[7]?.querySelector("input")?.value || "", // Pega o valor atual (pode ser o auto-gerado)
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
    aplicarIndentacao(row); // Re-aplica a indentação após preencher o NÍVEL
    cells[3].querySelector("select").value = data.SITE || "1";
    cells[4].querySelector("select").value = data.ALTERNATIVA || "*";
    cells[5].querySelector("input").value = (data.CODIGO_MATERIAL || "").toUpperCase();
    cells[6].querySelector("select").value = data.TIPO_ESTRUTURA || "Manufatura";
    // A coluna LINHA é preenchida por `atualizarColunaLinha()`, não diretamente aqui.
    cells[8].querySelector("input").value = (data.ITEM_COMPONENTE || "").toUpperCase();
    // Substitui vírgula por ponto para QTDE_MONTAGEM e remove "0" se não for o único caractere
    cells[9].querySelector("input").value = (data.QTDE_MONTAGEM === "0" ? "" : data.QTDE_MONTAGEM || "").replace(",", ".");
    // AQUI: Garantir que UNIDADE_MEDIDA seja preenchida corretamente
    cells[10].querySelector("input").value = (data.UNIDADE_MEDIDA || "").toLowerCase(); // Garante minúsculas
    cells[11].querySelector("select").value = data.FATOR_SUCATA || "0";
}

/**
 * Verifica e destaca linhas que contêm dados duplicados na tabela.
 * Ignora linhas completamente vazias e o destaque pode ser desativado via checkbox.
 */
function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    const hashes = new Map(); // Mapa para armazenar hashes e as linhas correspondentes
    linhas.forEach(row => { row.classList.remove("highlight-duplicate"); }); // Remove destaque anterior
    if (ignorarDuplicatas) return; // Sai se a opção de ignorar estiver ativa

    linhas.forEach((tr) => {
        const data = getLinhaData(tr);
        // Ignora linhas que são completamente vazias (ambos CODIGO_MATERIAL e ITEM_COMPONENTE vazios)
        if (data.CODIGO_MATERIAL === "" && data.ITEM_COMPONENTE === "") return;

        // Cria um hash único para identificar duplicatas
        // Adicionando NIVEL e TIPO_ESTRUTURA ao hash para uma verificação mais precisa de duplicatas
        const hash = `${data.SITE}|${data.ALTERNATIVA}|${data.CODIGO_MATERIAL}|${data.NIVEL}|${data.TIPO_ESTRUTURA}|${data.LINHA}|${data.ITEM_COMPONENTE}|${data.UNIDADE_MEDIDA}|${data.FATOR_SUCATA}`;
        if (!hashes.has(hash)) hashes.set(hash, []);
        hashes.get(hash).push(tr); // Adiciona a linha ao array do hash correspondente
    });

    // Percorre o mapa para encontrar e destacar hashes com mais de uma linha (duplicatas)
    for (const [hash, rows] of hashes) {
        if (rows.length > 1) {
            rows.forEach(row => { row.classList.add("highlight-duplicate"); }); // Adiciona classe de destaque
        }
    }
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
    delete parts[0]; // Remove o primeiro elemento que é a string completa da regex
    for (let i = 1; i <= 3; i++) {
        parts[i] = parseInt(parts[i]).toString(16); // Converte para hexadecimal
        if (parts[i].length === 1) parts[i] = "0" + parts[i]; // Adiciona zero à esquerda se necessário
    }
    return "#" + parts.join("").toUpperCase(); // Junta as partes e retorna em maiúsculas
}

/**
 * Exporta os dados da tabela atual para um arquivo Excel (.xlsx).
 * Inclui apenas as colunas especificadas e ignora linhas "vazias" (separadores).
 */
function exportarParaExcel() {
    // Define os cabeçalhos desejados para a exportação
    const ws_data = [
        ["NIVEL", "SITE", "ALTERNATIVA", "CODIGO_MATERIAL", "TIPO ESTRUTURA", "LINHA", "ITEM_COMPONENTE", "QTDE_MONTAGEM", "UNIDADE DE MEDIDA", "FATOR_SUCATA"]
    ];

    tabela.querySelectorAll("tr").forEach(row => {
        const rowData = getLinhaData(row);

        // Filtra linhas que são "separadores" (ambos vazios)
        if (rowData.CODIGO_MATERIAL === "" && rowData.ITEM_COMPONENTE === "") {
            return; // Pula esta linha, não a inclui na exportação
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

    if (ws_data.length <= 1) { // Apenas os cabeçalhos estão presentes
        Swal.fire("ℹ️ Nada para Exportar", "A tabela está vazia ou contém apenas linhas sem dados preenchidos.", "info");
        return;
    }

    const ws = XLSX.utils.aoa_to_sheet(ws_data); // Cria a planilha a partir do array de arrays
    const wb = XLSX.utils.book_new(); // Cria um novo livro Excel
    XLSX.utils.book_append_sheet(wb, ws, "Lista Tecnica"); // Adiciona a planilha ao livro
    XLSX.writeFile(wb, "Lista_Tecnica.xlsx"); // Escreve e baixa o arquivo Excel

    Swal.fire("✅ Exportado!", "A lista foi exportada para 'Lista_Tecnica.xlsx'.", "success");
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
        // Converte a planilha para um array de arrays, onde a primeira linha é o cabeçalho
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (json.length === 0 || !json[0]) { // Adicionado !json[0] para verificar se há cabeçalhos
            Swal.fire("⚠️ Arquivo Vazio ou Inválido", "O arquivo Excel está vazio, não contém cabeçalhos ou dados.", "warning");
            return;
        }

        // Limpa a tabela existente antes de carregar novos dados
        tabela.innerHTML = "";

        // Mapeia os cabeçalhos do Excel para encontrar os índices das colunas
        // Normaliza os cabeçalhos do Excel: remove espaços, substitui por '_' e coloca em maiúsculas
        const headers = json[0].map(h => String(h).trim().replace(/\s/g, '_').toUpperCase());
        const dataRows = json.slice(1); // Remove a linha de cabeçalho dos dados

        /**
         * Mapeamento dos índices das colunas esperadas com base nos cabeçalhos do Excel normalizados.
         * Nomes corrigidos para corresponder ao padrão do Excel após normalização.
         */
        const colIndices = {
            NIVEL: headers.indexOf("NIVEL"),
            SITE: headers.indexOf("SITE"),
            ALTERNATIVA: headers.indexOf("ALTERNATIVA"),
            CODIGO_MATERIAL: headers.indexOf("CODIGO_MATERIAL"),
            TIPO_ESTRUTURA: headers.indexOf("TIPO_ESTRUTURA"),
            LINHA: headers.indexOf("LINHA"), // Incluído para referência, embora seja autopreenchido
            ITEM_COMPONENTE: headers.indexOf("ITEM_COMPONENTE"),
            QTDE_MONTAGEM: headers.indexOf("QTDE_MONTAGEM"),
            // Múltiplas opções para UNIDADE_MEDIDA para maior robustez
            UNIDADE_MEDIDA: headers.indexOf("UNIDADE_DE_MEDIDA") !== -1 ?
                            headers.indexOf("UNIDADE_DE_MEDIDA") :
                            headers.indexOf("UNIDADE_MEDIDA"), // Tenta sem underscore se o primeiro falhar
            FATOR_SUCATA: headers.indexOf("FATOR_SUCATA")
        };

        // Log de depuração para verificar mapeamento de colunas (pode ser removido em produção)
        console.log("Cabeçalhos do Excel (normalizados):", headers);
        console.log("Índices de Colunas Mapeados:", colIndices);
        for (const key in colIndices) {
            if (colIndices[key] === -1) {
                console.warn(`Atenção: A coluna "${key}" não foi encontrada no Excel. Verifique o nome do cabeçalho.`);
            }
        }

        // Processa cada linha de dados do Excel
        dataRows.forEach(rowData => {
            const rowObj = {
                // Usa String() para garantir que os valores sejam tratados como string
                // Garante valor vazio se o índice for -1 ou o dado for undefined/null
                NIVEL: colIndices.NIVEL !== -1 && rowData[colIndices.NIVEL] !== undefined ? String(rowData[colIndices.NIVEL]) : "",
                SITE: colIndices.SITE !== -1 && rowData[colIndices.SITE] !== undefined ? String(rowData[colIndices.SITE]) : "1",
                ALTERNATIVA: colIndices.ALTERNATIVA !== -1 && rowData[colIndices.ALTERNATIVA] !== undefined ? String(rowData[colIndices.ALTERNATIVA]) : "*",
                CODIGO_MATERIAL: colIndices.CODIGO_MATERIAL !== -1 && rowData[colIndices.CODIGO_MATERIAL] !== undefined ? String(rowData[colIndices.CODIGO_MATERIAL]) : "",
                TIPO_ESTRUTURA: colIndices.TIPO_ESTRUTURA !== -1 && rowData[colIndices.TIPO_ESTRUTURA] !== undefined ? String(rowData[colIndices.TIPO_ESTRUTURA]) : "Manufatura",
                ITEM_COMPONENTE: colIndices.ITEM_COMPONENTE !== -1 && rowData[colIndices.ITEM_COMPONENTE] !== undefined ? String(rowData[colIndices.ITEM_COMPONENTE]) : "",
                QTDE_MONTAGEM: colIndices.QTDE_MONTAGEM !== -1 && rowData[colIndices.QTDE_MONTAGEM] !== undefined ? String(rowData[colIndices.QTDE_MONTAGEM]) : "",
                // AQUI: Garantir que UNIDADE_MEDIDA seja corretamente passada e em minúsculas
                UNIDADE_MEDIDA: colIndices.UNIDADE_MEDIDA !== -1 && rowData[colIndices.UNIDADE_MEDIDA] !== undefined ? String(rowData[colIndices.UNIDADE_MEDIDA]).toLowerCase() : "",
                FATOR_SUCATA: colIndices.FATOR_SUCATA !== -1 && rowData[colIndices.FATOR_SUCATA] !== undefined ? String(rowData[colIndices.FATOR_SUCATA]) : "0"
            };
            const newRow = criarLinha(rowObj);
            tabela.appendChild(newRow);
        });

        // Executa as ações de atualização após a importação (SEQ, LINHA, Duplicatas)
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
    atualizarColunaLinha(); // Garante que LINHA seja preenchida corretamente para todas as linhas
    verificarDuplicatas();
}

// --- Event Listeners para Botões da Barra de Ferramentas ---

/**
 * Listener para o botão "Criar Nova Lista".
 * Limpa a tabela e adiciona 10 linhas vazias.
 */
document.getElementById("criarListaBtn").addEventListener("click", () => {
    tabela.innerHTML = ""; // Limpa a tabela
    criar10Linhas(); // Adiciona 10 linhas vazias
    acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas
    Swal.fire("✅ Lista Criada!", "10 novas linhas foram adicionadas.", "success");
});

/**
 * Listener para o botão "Continuar Lista".
 * Adiciona 10 linhas vazias ao final da tabela existente.
 */
document.getElementById("continuarListaBtn").addEventListener("click", () => {
    criar10Linhas(); // Adiciona 10 linhas vazias
    acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas
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
    cacheCopiado = linhasSelecionadas.map(row => getLinhaData(row)); // Armazena os dados das linhas copiadas
    Swal.fire("✅ Copiado!", `${cacheCopiado.length} linhas copiadas.`, "success");
});

/**
 * Listener para o botão "Colar".
 * Cola os dados do cache em novas linhas ou sobrescreve linhas selecionadas.
 */
document.getElementById("colarBtn").addEventListener("click", () => {
    if (cacheCopiado.length === 0) {
        Swal.fire("ℹ️ Nada para Colar", "Nenhum dado copiado.", "info");
        return;
    }

    const linhasSelecionadas = Array.from(tabela.querySelectorAll(".linha-selecao:checked")).map(cb => cb.closest("tr"));
    // Determina o índice de início da colagem: primeira linha selecionada ou final da tabela
    const startIndex = linhasSelecionadas.length > 0 ? Array.from(tabela.rows).indexOf(linhasSelecionadas[0]) : tabela.rows.length;

    cacheCopiado.forEach((rowData, i) => {
        const targetRow = tabela.rows[startIndex + i];
        if (targetRow) {
            // Se houver uma linha alvo, preenche ela
            preencherLinha(targetRow, rowData);
        } else {
            // Se não houver, cria uma nova linha
            const newRow = criarLinha(rowData);
            tabela.appendChild(newRow);
        }
    });

    acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas após colar
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

    // Confirmação antes de deletar
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
    tabela.insertBefore(novaLinha, primeiraLinhaSelecionada); // Insere antes da linha selecionada
    acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas
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
        // Se houver uma próxima linha, insere antes dela
        tabela.insertBefore(novaLinha, ultimaLinhaSelecionada.nextElementSibling);
    } else {
        // Caso contrário, adiciona ao final da tabela
        tabela.appendChild(novaLinha);
    }
    acaoImportouOuAdicionouLinhas(); // Atualiza SEQ, LINHA, e Duplicatas
    Swal.fire("⬇️ Inserido", "Uma nova linha foi inserida abaixo da seleção.", "success");
});

/**
 * Listener para o botão "Alternar SEQ".
 * Alterna a visibilidade da coluna "SEQ".
 */
document.getElementById("toggleSeqBtn").addEventListener("click", () => {
    seqAtivo = !seqAtivo; // Inverte o estado
    document.getElementById("listaTabela").classList.toggle("seq-col-hidden", !seqAtivo); // Adiciona/remove classe CSS
    document.getElementById("toggleSeqBtn").textContent = seqAtivo ? "👁️ SEQ" : "✖️ SEQ"; // Atualiza texto do botão
});

/**
 * Listener para o botão "Alternar NÍVEL".
 * Alterna a visibilidade da coluna "NÍVEL".
 */
document.getElementById("toggleNivelColBtn").addEventListener("click", () => {
    nivelColVisivel = !nivelColVisivel; // Inverte o estado
    document.getElementById("listaTabela").classList.toggle("nivel-col-hidden", !nivelColVisivel); // Adiciona/remove classe CSS
    document.getElementById("toggleNivelColBtn").textContent = nivelColVisivel ? "👁️ NÍVEL" : "✖️ NÍVEL"; // Atualiza texto do botão
});

/**
 * Listener para o botão "Alternar Régua".
 * Alterna o efeito de destaque de linha ao passar o mouse (régua).
 */
document.getElementById("toggleHoverEffectBtn").addEventListener("click", () => {
    hoverEffectAtivo = !hoverEffectAtivo; // Inverte o estado
    const tableElement = document.getElementById("listaTabela");
    tableElement.classList.toggle("hover-effect", hoverEffectAtivo);
    tableElement.classList.toggle("no-hover-effect", !hoverEffectAtivo);
    document.getElementById("toggleHoverEffectBtn").textContent = hoverEffectAtivo ? "📏 Régua" : "✖️ Régua"; // Atualiza texto do botão
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
    if (demarcarLinha) removerDemarcacao = false; // Desativa "Remover demarcação" se "Demarcar linha" está ativo
    document.getElementById("removerDemarcacaoCheckbox").checked = false;
});

/**
 * Listener para o checkbox "Remover demarcação".
 * Ativa/desativa o modo de remoção de cor ao clicar.
 */
document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", (e) => {
    removerDemarcacao = e.target.checked;
    if (removerDemarcacao) demarcarLinha = false; // Desativa "Demarcar linha" se "Remover demarcação" está ativo
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
    button.style.color = getContrastColor(color); // Define cor do texto para contraste
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
        return "black"; // Retorna preto para cores não-hexadecimais
    }
    const r = parseInt(hexcolor.slice(1, 3), 16);
    const g = parseInt(hexcolor.slice(3, 5), 16);
    const b = parseInt(hexcolor.slice(5, 7), 16);
    // Fórmula HSP (Highly Sensitive Poo) para luminância
    const hsp = Math.sqrt(
        0.299 * (r * r) +
        0.587 * (g * g) +
        0.114 * (b * b)
    );
    return (hsp > 127.5) ? "black" : "white"; // Retorna preto para cores claras, branco para escuras
}

/**
 * Listener para o checkbox "Ignorar destaque de duplicatas".
 * Ativa/desativa a funcionalidade de destaque de duplicatas e atualiza a tabela.
 */
document.getElementById("ignorarDuplicatasCheckbox").addEventListener("change", (e) => {
    ignorarDuplicatas = e.target.checked;
    verificarDuplicatas(); // Re-executa para remover/adicionar destaque
});

// --- Sincronização de Checkboxes de Seleção Total ---

/**
 * Listener para o checkbox "Marcar/Desmarcar Todos" no painel de controle.
 * Sincroniza o estado com o checkbox do cabeçalho da tabela e marca/desmarca todas as linhas.
 */
document.getElementById("toggleAllCheckboxesHeader").addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    document.getElementById("toggleAllCheckboxes").checked = isChecked; // Sincroniza com o checkbox do cabeçalho da tabela
    tabela.querySelectorAll(".linha-selecao").forEach(checkbox => {
        checkbox.checked = isChecked; // Marca/desmarca todas as linhas
    });
});

/**
 * Listener para o checkbox "Marcar/Desmarcar Todos" no cabeçalho da tabela.
 * Sincroniza o estado com o checkbox do painel de controle e marca/desmarca todas as linhas.
 */
document.getElementById("toggleAllCheckboxes").addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    document.getElementById("toggleAllCheckboxesHeader").checked = isChecked; // Sincroniza com o checkbox do painel de controle
    tabela.querySelectorAll(".linha-selecao").forEach(checkbox => {
        checkbox.checked = isChecked; // Marca/desmarca todas as linhas
    });
});

// --- Listener para Input de Arquivo (Importação) ---

/**
 * Listener para o input de arquivo (botão "Importar Lista").
 * Dispara a função `carregarExcel` quando um arquivo é selecionado.
 */
document.getElementById("inputFile").addEventListener("change", function (e) {
    carregarExcel(e.target);
    e.target.value = ''; // Limpa o input para que o mesmo arquivo possa ser selecionado novamente
});

// --- Inicialização da Aplicação ---

/**
 * Executado quando o DOM está completamente carregado.
 * Inicializa a tabela com 10 linhas se estiver vazia, ou atualiza o conteúdo existente.
 * Define o estado inicial de visibilidade das colunas e efeitos.
 */
document.addEventListener("DOMContentLoaded", () => {
    if (tabela.rows.length === 0) {
        criar10Linhas(); // Cria 10 linhas iniciais se a tabela estiver vazia
        acaoImportouOuAdicionouLinhas(); // Aplica regras de SEQ, LINHA, e Duplicatas
        Swal.fire("🎉 Bem-vindo!", "A lista foi inicializada com 10 linhas para você começar.", "info");
    } else {
        // Se já houver conteúdo (e.g., após um refresh ou carregamento prévio),
        // garante que LINHA e SEQ estejam corretas.
        acaoImportouOuAdicionouLinhas();
    }

    // Aplica classes CSS para controlar a visibilidade das colunas e efeitos
    const tabelaElement = document.getElementById("listaTabela");
    if (!seqAtivo) tabelaElement.classList.add("seq-col-hidden");
    if (!nivelColVisivel) tabelaElement.classList.add("nivel-col-hidden");
    if (hoverEffectAtivo) tabelaElement.classList.add("hover-effect");
    else tabelaElement.classList.add("no-hover-effect");

    // Inicializa o texto dos botões de alternância de visibilidade
    document.getElementById("toggleSeqBtn").textContent = seqAtivo ? "👁️ SEQ" : "✖️ SEQ";
    document.getElementById("toggleNivelColBtn").textContent = nivelColVisivel ? "👁️ NÍVEL" : "✖️ NÍVEL";
    document.getElementById("toggleHoverEffectBtn").textContent = hoverEffectAtivo ? "📏 Régua" : "✖️ Régua";
});