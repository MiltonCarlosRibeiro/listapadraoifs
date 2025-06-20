let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true; // SEQ estará visível por padrão
let nivelColVisivel = true; // Nível estará visível por padrão
let corSelecionada = ""; // Cor hex (#RRGGBB) para aplicar
let demarcarLinha = false; // true: clica na linha, false: clica na célula
let removerDemarcacao = false; // true: cliques removem cores, false: cliques aplicam cores
let ignorarDuplicatas = false; // Nova variável para ignorar destaque de duplicatas
let hoverEffectAtivo = true; // Estado inicial da régua do mouse

// Definição das cores para os níveis (ATUALIZADAS - CORRIGIDO NÍVEL 3)
const nivelColors = [
    "#4664cf", // Nível 1 - Azul anil
    "#CD5C5C", // Nível 2 - Vermelho Indiano
    "#B3E6B3", // Nível 3 - Verde Pastel (derivado do azul original do Nível 1, com tom abaixado para pastel)
    "#FFD700", // Nível 4 - Ouro
    "#8A2BE2", // Nível 5 - Azul Violeta
    "#FF8C00", // Nível 6 - Laranja Escuro
    "#00CED1", // Nível 7 - Turquesa Escuro
    "#FF69B4", // Nível 8 - Rosa Choque
    "#9ACD32", // Nível 9 - Verde Amarelado
    "#DA70D6"  // Nível 10 - Orquídea
];

const unidades = ["un", "cj", "kg", "mm", "m"];
const tiposEstrutura = ["Manufatura", "Comprado", ""];
const fatorSucata = ["0", "15", ""];
const linhaValores = Array.from({ length: 80 }, (_, i) => String((i + 1) * 10));
const alternativas = ["*", ""];
const siteValores = ["1", ""];
const niveis = Array.from({ length: 10 }, (_, i) => (i + 1).toString());

function inputCell(type, readOnly = false, value = "", isPasteTarget = false, className = "") {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = type;
    input.readOnly = readOnly;
    input.value = (value || "").toUpperCase();
    if (className) td.classList.add(className); // Adiciona a classe à TD para CSS de nível

    input.addEventListener("input", (e) => {
        e.target.value = e.target.value.toUpperCase();
        verificarDuplicatas();
        if (td.classList.contains('nivel-col')) { // Se for a coluna de nível, aplica indentação
            aplicarIndentacao(e.target.closest('tr'));
            // Dispara um evento personalizado para notificar o filtro sobre a mudança
            e.target.dispatchEvent(new Event('change', { bubbles: true }));
        }
    });

    // === Nova lógica para navegação com Enter ===
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault(); // Impede o comportamento padrão do Enter (ex: submit de formulário)

            const currentInput = e.target;
            const currentTd = currentInput.closest('td');
            const currentRow = currentInput.closest('tr');
            const currentRowIndex = Array.from(tabela.rows).indexOf(currentRow);
            const currentCellIndex = Array.from(currentRow.children).indexOf(currentTd);

            // Tenta focar no input da mesma coluna na próxima linha
            const nextRow = tabela.rows[currentRowIndex + 1];
            if (nextRow) {
                const nextTd = nextRow.children[currentCellIndex];
                const nextInput = nextTd?.querySelector('input, select');
                if (nextInput) {
                    nextInput.focus();
                } else {
                    // Se não encontrar input na mesma coluna na próxima linha, tenta a próxima célula na linha atual
                    const nextCellInRow = currentRow.children[currentCellIndex + 1];
                    const nextInputInRow = nextCellInRow?.querySelector('input, select');
                    if (nextInputInRow) {
                        nextInputInRow.focus();
                    }
                }
            } else {
                // Se for a última linha, cria uma nova linha e foca no primeiro input
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                atualizarSequencias();
                verificarDuplicatas();

                // Foca no input da mesma coluna na nova linha
                const firstInputInNewRow = newRow.children[currentCellIndex]?.querySelector('input, select');
                if (firstInputInNewRow) {
                    firstInputInNewRow.focus();
                }
            }
        }
    });
    // ===========================================

    // Evento de clique para pintura de célula (com nova lógica)
    input.addEventListener("click", (e) => {
        if (!demarcarLinha) { // Se não for para demarcar a linha inteira
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
    select.addEventListener("change", verificarDuplicatas);

    // Evento de clique para pintura de célula (com nova lógica)
    select.addEventListener("click", (e) => {
        if (!demarcarLinha) { // Se não for para demarcar a linha inteira
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

    // === Nova lógica para navegação com Enter em selects ===
    select.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault(); // Impede o comportamento padrão do Enter

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
                    // Se não encontrar, tenta a próxima célula na linha atual
                    const nextCellInRow = currentRow.children[currentCellIndex + 1];
                    const nextInputInRow = nextCellInRow?.querySelector('input, select');
                    if (nextInputInRow) {
                        nextInputInRow.focus();
                    }
                }
            } else {
                const newRow = criarLinhaVazia();
                tabela.appendChild(newRow);
                atualizarSequencias();
                verificarDuplicatas();

                const firstInputInNewRow = newRow.children[currentCellIndex]?.querySelector('input, select');
                if (firstInputInNewRow) {
                    firstInputInNewRow.focus();
                }
            }
        }
    });
    // ===========================================

    td.appendChild(select);
    return td;
}

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
        if (e.target.tagName.match(/INPUT|SELECT|BUTTON/)) {
            return;
        }

        if (demarcarLinha) {
            if (removerDemarcacao) {
                row.style.backgroundColor = "";
            } else if (corSelecionada) {
                // Se a linha já tem a cor selecionada, remove
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
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || "", true));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));
    row.appendChild(selectCell(linhaValores, v.LINHA || "10"));
    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || "", true));
    row.appendChild(inputCell("text", false, v.QTDE_MONTAGEM || "0", true));
    row.appendChild(selectCell(unidades, v.UNIDADE_MEDIDA || "un"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    aplicarIndentacao(row);

    // Garante que as colunas de SEQ e Nível respeitem o estado inicial
    if (!seqAtivo) {
        seqTd.style.display = "none";
    }
    if (!nivelColVisivel) {
        nivelCell.style.display = "none";
    }

    return row;
}

function criarLinhaVazia() {
    return criarLinha({});
}

function atualizarSequencias() {
    const linhas = tabela.querySelectorAll("tr");
    linhas.forEach((row, index) => {
        const seqTd = row.querySelectorAll("td")[1];
        if (seqTd) {
            seqTd.textContent = (index + 1) * 10;
        }
    });
}

function aplicarIndentacao(row) {
    // Remove todas as classes de nível existentes
    for (let i = 1; i <= 10; i++) {
        row.classList.remove(`nivel-${i}`);
    }
    const nivelInput = row.querySelectorAll("td")[2]?.querySelector("input");
    if (nivelInput) {
        let nivel = parseInt(nivelInput.value);
        if (!isNaN(nivel) && nivel >= 1 && nivel <= 10) {
            row.classList.add(`nivel-${nivel}`);
        }
    }
}

function criar10Linhas() {
    for (let i = 0; i < 10; i++) {
        const novaLinha = criarLinhaVazia();
        tabela.appendChild(novaLinha);
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
        LINHA: cells[7]?.querySelector("select")?.value || "",
        ITEM_COMPONENTE: cells[8]?.querySelector("input")?.value.trim().toUpperCase() || "",
        QTDE_MONTAGEM: cells[9]?.querySelector("input")?.value.trim() || "",
        UNIDADE_MEDIDA: cells[10]?.querySelector("select")?.value || "",
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
    cells[7].querySelector("select").value = data.LINHA || "10";
    cells[8].querySelector("input").value = (data.ITEM_COMPONENTE || "").toUpperCase();
    cells[9].querySelector("input").value = (data.QTDE_MONTAGEM || "0").replace(",", ".");
    cells[10].querySelector("select").value = data.UNIDADE_MEDIDA || "un";
    cells[11].querySelector("select").value = data.FATOR_SUCATA || "0";
}

function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    const hashes = new Map();
    // Cores de atenção e de nível para evitar sobrescrita
    const protectedColors = new Set([
        "#FFBF00", "#28A745", "#DC3545", // Cores de atenção
        ...nivelColors // Cores de nível
    ].map(c => c.toUpperCase()));

    // Primeiro, remove a cor de duplicidade de todas as linhas
    linhas.forEach(row => {
        const currentColor = rgbToHex(row.style.backgroundColor);
        if (currentColor === "#F0E6FF") { // Cor roxa clara de duplicidade
            row.style.backgroundColor = "";
        }
    });

    if (ignorarDuplicatas) {
        return;
    }

    // Em seguida, verifica e aplica a cor de duplicidade
    linhas.forEach((tr) => {
        const data = getLinhaData(tr);

        if (data.CODIGO_MATERIAL === "" && data.ITEM_COMPONENTE === "") return;

        const hash = `${data.SITE}|${data.ALTERNATIVA}|${data.CODIGO_MATERIAL}|${data.NIVEL}|${data.TIPO_ESTRUTURA}|${data.LINHA}|${data.ITEM_COMPONENTE}|${data.QTDE_MONTAGEM}|${data.UNIDADE_MEDIDA}|${data.FATOR_SUCATA}`;

        if (!hashes.has(hash)) hashes.set(hash, []);
        hashes.get(hash).push(tr);
    });

    for (const [hash, rows] of hashes) {
        if (rows.length > 1) {
            rows.forEach(row => {
                const currentColor = rgbToHex(row.style.backgroundColor);
                // Aplica a cor de duplicidade APENAS se a linha não tiver uma cor protegida
                if (!protectedColors.has(currentColor)) {
                    row.style.backgroundColor = "#f0e6ff"; // Cor roxa clara
                }
            });
        }
    }
}


function rgbToHex(rgb) {
    if (!rgb || rgb.indexOf('rgb') === -1) {
        return rgb ? rgb.toUpperCase() : "";
    }
    const parts = rgb.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
    if (!parts) return "";
    delete parts[0];
    for (let i = 1; i <= 3; i++) {
        parts[i] = parseInt(parts[i]).toString(16);
        if (parts[i].length === 1) parts[i] = "0" + parts[i];
    }
    return "#" + parts.join("").toUpperCase();
}


// EVENTOS
document.getElementById("criarListaBtn").addEventListener("click", () => {
    tabela.innerHTML = "";
    criar10Linhas();
    Swal.fire("✅ Lista Criada!", "10 novas linhas foram adicionadas.", "success");
    atualizarSequencias();
    verificarDuplicatas();
});

document.getElementById("continuarListaBtn").addEventListener("click", () => {
    criar10Linhas();
    Swal.fire("➕ Adicionado", "10 novas linhas foram inseridas.", "success");
    atualizarSequencias();
    verificarDuplicatas();
});

document.getElementById("deletarSelecionadosBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length === 0) {
        Swal.fire("⚠️ Nada selecionado", "Marque uma linha para deletar.", "info");
        return;
    }
    Swal.fire({
        title: "Tem certeza?",
        text: `Você vai deletar ${selecionados.length} linha(s).`,
        icon: "warning",
        showCancelButton: true,
        confirmButtonColor: "#3085d6",
        cancelButtonColor: "#d33",
        confirmButtonText: "Sim, deletar!",
        cancelButtonText: "Cancelar"
    }).then((result) => {
        if (result.isConfirmed) {
            selecionados.forEach(cb => cb.closest("tr").remove());
            atualizarSequencias();
            verificarDuplicatas();
            Swal.fire("🗑️ Removido", "Linhas selecionadas foram deletadas.", "success");
        }
    });
});

document.getElementById("inserirAcimaBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length !== 1) return Swal.fire("⚠️ Selecione uma única linha", "Para inserir acima.", "info");

    const row = selecionados[0].closest("tr");
    const nova = criarLinhaVazia();
    tabela.insertBefore(nova, row);
    atualizarSequencias();
    verificarDuplicatas();
    Swal.fire("⬆️ Linha inserida", "Nova linha foi adicionada acima.", "success");
});

document.getElementById("inserirAbaixoBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length !== 1) return Swal.fire("⚠️ Selecione uma única linha", "Para inserir abaixo.", "info");

    const row = selecionados[0].closest("tr");
    const nova = criarLinhaVazia();
    tabela.insertBefore(nova, row.nextSibling);
    atualizarSequencias();
    verificarDuplicatas();
    Swal.fire("⬇️ Linha inserida", "Nova linha foi adicionada abaixo.", "success");
});

document.getElementById("copiarSelecionadoBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");

    if (selecionados.length === 0) {
        Swal.fire("⚠️ Nada selecionado", "Marque uma linha.", "info");
        return;
    }

    cacheCopiado = [];
    cacheCopiado = Array.from(selecionados).map(cb => getLinhaData(cb.closest("tr")));
    Swal.fire("📋 Copiado", `${cacheCopiado.length} linha(s) copiadas.`, "success");
});


document.getElementById("colarBtn").addEventListener("click", () => {
    if (cacheCopiado.length === 0) {
        Swal.fire("⚠️ Nada para colar", "Copie algo primeiro.", "info");
        return;
    }

    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length !== 1) {
        Swal.fire("⚠️ Selecione uma linha", "Use uma linha como referência para colar.", "info");
        return;
    }

    const trBase = selecionados[0].closest("tr");
    let index = Array.from(tabela.rows).indexOf(trBase);

    cacheCopiado.forEach((linhaData, i) => {
        let targetRow = tabela.rows[index + i];
        if (!targetRow) {
            const novaLinha = criarLinhaVazia();
            tabela.appendChild(novaLinha);
            targetRow = novaLinha;
        }
        preencherLinha(targetRow, linhaData);
    });

    atualizarSequencias();
    verificarDuplicatas();
    Swal.fire("📥 Colado", "Conteúdo copiado foi inserido com sucesso.", "success");
});

document.getElementById("salvarListaBtn").addEventListener("click", () => {
    const dados = Array.from(tabela.rows).map(getLinhaData);
    if (dados.length === 0) {
        Swal.fire("⚠️ Lista vazia", "Não há dados para exportar.", "info");
        return;
    }
    const ws = XLSX.utils.json_to_sheet(dados);

    // Ajustes de largura de coluna para o Excel
    ws['!cols'] = ws['!cols'] || [];
    // Adiciona Largura para as colunas: Nível (index 2) e SEQ (index 1)
    ws['!cols'][2] = { wch: 5 }; // Nível
    ws['!cols'][1] = { wch: 5 }; // SEQ
    // Adiciona Largura para Código Material e Item Componente
    ws['!cols'][5] = { wch: 20 }; // CODIGO_MATERIAL
    ws['!cols'][8] = { wch: 20 }; // ITEM_COMPONENTE

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ListaIFS");
    XLSX.writeFile(wb, `lista_padrao_ifs_${new Date().toISOString().split("T")[0]}.xlsx`);
    Swal.fire("💾 Lista Exportada", "Arquivo Excel gerado com sucesso.", "success");
});

document.getElementById("inputFile").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    Swal.fire({
        title: "Carregando lista...",
        text: "Por favor, aguarde.",
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        tabela.innerHTML = "";

        json.forEach(rowData => {
            const mappedData = {
                NIVEL: String(rowData.NIVEL || rowData.Nivel || ""),
                SITE: rowData.SITE || rowData.Site || "1",
                ALTERNATIVA: rowData.ALTERNATIVA || rowData.Alternativa || "*",
                CODIGO_MATERIAL: rowData["CÓDIGO_MATERIAL"] || rowData["CODIGO_MATERIAL"] || rowData["CODIGO MATERIAL"] || rowData.Codigo_Material || rowData.CodigoMaterial || "",
                TIPO_ESTRUTURA: rowData["TIPO ESTRUTURA"] || rowData.Tipo_Estrutura || rowData.TipoEstrutura || "Manufatura",
                LINHA: String(rowData.LINHA || rowData.Linha || "10"),
                ITEM_COMPONENTE: rowData.ITEM_COMPONENTE || rowData.Item_Componente || rowData.ItemComponente || "",
                QTDE_MONTAGEM: String(rowData.QTDE_MONTAGEM || rowData.Qtde_Montagem || rowData.QtdeMontagem || "0"),
                UNIDADE_MEDIDA: rowData["UNIDADE DE MEDIDA"] || rowData.Unidade_Medida || rowData.UnidadeMedida || "un",
                FATOR_SUCATA: String(rowData.FATOR_SUCATA || rowData.Fator_Sucata || rowData.FatorSucata || "0")
            };
            const novaLinha = criarLinha(mappedData);
            tabela.appendChild(novaLinha);
        });

        atualizarSequencias();
        verificarDuplicatas();
        Swal.fire("✅ Lista carregada", "A tabela foi preenchida com sucesso.", "success");
    } catch (error) {
        console.error("Erro ao carregar o arquivo:", error);
        Swal.fire("❌ Erro", "Não foi possível carregar a lista. Verifique o formato do arquivo ou o console para detalhes.", "error");
    } finally {
        e.target.value = '';
    }
});

const nivelColorButtonsContainer = document.getElementById("nivelColorButtons");
// Limpa os botões existentes antes de recriar
nivelColorButtonsContainer.innerHTML = '';
nivelColors.forEach((color, index) => {
    const button = document.createElement("button");
    button.className = `paint-btn btn-nivel-${index + 1}`;
    button.dataset.color = color;
    button.textContent = index + 1;
    button.style.backgroundColor = color; // Aplica a cor diretamente ao background do botão
    button.style.color = getContrastYIQ(color); // Define a cor do texto para contraste
    button.addEventListener("click", function() {
        corSelecionada = this.dataset.color;
        document.getElementById("removerDemarcacaoCheckbox").checked = false;
        removerDemarcacao = false;
        Swal.fire({
            title: `Cor de Nível ${index + 1} selecionada!`,
            text: `Clique na ${demarcarLinha ? 'linha' : 'célula'} que deseja pintar.`,
            icon: "info",
            timer: 2000,
            showConfirmButton: false
        });
    });
    nivelColorButtonsContainer.appendChild(button);
});

// Cores de Atenção
document.querySelectorAll("#attentionColorButtons .paint-btn").forEach(button => {
    button.addEventListener("click", function() {
        corSelecionada = this.dataset.color;
        document.getElementById("removerDemarcacaoCheckbox").checked = false;
        removerDemarcacao = false;
        const labelText = this.textContent; // Pega o texto do botão
        Swal.fire({
            title: `Cor "${labelText}" selecionada!`,
            text: `Clique na ${demarcarLinha ? 'linha' : 'célula'} que deseja pintar.`,
            icon: "info",
            timer: 2000,
            showConfirmButton: false
        });
    });
});

document.getElementById("demarcarLinhaCheckbox").addEventListener("change", function() {
    demarcarLinha = this.checked;
    if (demarcarLinha) {
        Swal.fire("Modo 'Demarcar linha' ativado!", "Agora, cliques aplicarão a cor selecionada à linha inteira.", "info");
    } else {
        Swal.fire("Modo 'Demarcar linha' desativado!", "Cliques demarcarão células individuais.", "info");
    }
});

document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", function() {
    removerDemarcacao = this.checked;
    if (removerDemarcacao) {
        corSelecionada = ""; // Limpa a cor selecionada para o modo de remoção
        Swal.fire("Modo 'Remover demarcação' ativado!", "Agora, cliques removerão as demarcações existentes. Selecione um botão de cor para sair deste modo.", "warning");
    } else {
        Swal.fire("Modo 'Remover demarcação' desativado!", "Pode voltar a demarcar.", "info");
    }
});

document.getElementById("clearPaintBtn").addEventListener("click", () => {
    corSelecionada = "";
    document.getElementById("removerDemarcacaoCheckbox").checked = false;
    demarcarLinha = false;
    document.getElementById("demarcarLinhaCheckbox").checked = false;
    removerDemarcacao = false;
    // Limpa TODAS as demarcações ao clicar neste botão
    const allCells = tabela.querySelectorAll("td");
    const allRows = tabela.querySelectorAll("tr");
    allCells.forEach(cell => cell.style.backgroundColor = "");
    allRows.forEach(row => row.style.backgroundColor = "");

    Swal.fire("Seleção de cor limpa e demarcações removidas!", "Agora o clique não aplicará cores e todas as demarcações foram apagadas.", "info", 2500);
});

document.getElementById("ignorarDuplicatasCheckbox").addEventListener("change", function() {
    ignorarDuplicatas = this.checked;
    verificarDuplicatas();
    if (ignorarDuplicatas) {
        Swal.fire("Destaque de duplicatas desativado!", "As linhas duplicadas não serão mais destacadas com roxo claro.", "info", 2000);
    } else {
        Swal.fire("Destaque de duplicatas ativado!", "Linhas duplicadas serão destacadas com roxo claro.", "info", 2000);
    }
});

document.getElementById("toggleAllCheckboxesHeader").addEventListener("change", function() {
    const isChecked = this.checked;
    document.querySelectorAll(".linha-selecao").forEach(checkbox => {
        checkbox.checked = isChecked;
    });
});


document.getElementById("toggleSeqBtn").addEventListener("click", () => {
    const tabelaElement = document.getElementById("listaTabela");
    seqAtivo = !seqAtivo; // Inverte o estado

    if (seqAtivo) {
        tabelaElement.classList.remove("seq-col-hidden");
        Swal.fire("SEQ Visível!", "A coluna SEQ está visível.", "info", 1500);
    } else {
        tabelaElement.classList.add("seq-col-hidden");
        Swal.fire("SEQ Oculta!", "A coluna SEQ está oculta.", "info", 1500);
    }
});

document.getElementById("toggleNivelColBtn").addEventListener("click", () => {
    const tabelaElement = document.getElementById("listaTabela");
    nivelColVisivel = !nivelColVisivel; // Inverte o estado

    if (nivelColVisivel) {
        tabelaElement.classList.remove("nivel-col-hidden");
        Swal.fire("NÍVEL Visível!", "A coluna NÍVEL está visível.", "info", 1500);
    } else {
        tabelaElement.classList.add("nivel-col-hidden");
        Swal.fire("NÍVEL Oculta!", "A coluna NÍVEL está oculta.", "info", 1500);
    }
});

document.getElementById("toggleHoverEffectBtn").addEventListener("click", () => {
    const tabelaElement = document.getElementById("listaTabela");
    hoverEffectAtivo = !hoverEffectAtivo; // Inverte o estado da régua

    if (hoverEffectAtivo) {
        tabelaElement.classList.remove("no-hover-effect");
        tabelaElement.classList.add("hover-effect"); // Garante que a classe de hover está presente
        Swal.fire("Régua Ativada!", "O destaque da linha ao passar o mouse está ativo.", "info", 1500);
    } else {
        tabelaElement.classList.remove("hover-effect"); // Remove a classe de hover
        tabelaElement.classList.add("no-hover-effect");
        // Limpar qualquer régua ativa ao desativar
        document.querySelectorAll("#listaTabela tbody tr").forEach(row => {
            // Verifica se a cor de fundo é a cor da régua antes de limpar, para não afetar demarcações manuais
            if (rgbToHex(row.style.backgroundColor) === "#FFE0B2") { // Cor da régua: #ffe0b2
                row.style.backgroundColor = "";
            }
        });
        Swal.fire("Régua Desativada!", "O destaque da linha foi removido.", "info", 1500);
    }
});


async function handlePasteMultipleLines(event) {
    // Permite colar apenas em células de input onde o `isPasteTarget` é true
    const targetCell = event.target;
    const td = targetCell.closest("td");
    const tr = targetCell.closest("tr");
    const rowIndex = Array.from(tabela.rows).indexOf(tr);

    let columnIndex = -1;
    const allCellsInRow = tr.querySelectorAll("td");

    // Verifica se o targetCell é um input dentro de um td com uma classe específica
    if (targetCell === allCellsInRow[2]?.querySelector("input")) { // NÍVEL (index 2)
        columnIndex = 2;
    } else if (targetCell === allCellsInRow[5]?.querySelector("input")) { // CÓDIGO_MATERIAL (index 5)
        columnIndex = 5;
    } else if (targetCell === allCellsInRow[8]?.querySelector("input")) { // ITEM_COMPONENTE (index 8)
        columnIndex = 8;
    } else if (targetCell === allCellsInRow[9]?.querySelector("input")) { // QTDE_MONTAGEM (index 9)
        columnIndex = 9;
    } else {
        // Se não for uma das colunas permitidas para colagem, impede o evento padrão
        return;
    }

    // Se qualquer modo de demarcação ou cor estiver selecionado, impede a colagem
    if (corSelecionada || demarcarLinha || removerDemarcacao) {
        Swal.fire("Modo de Demarcação Ativo", "Desative o modo de demarcação ou limpe a cor selecionada para colar.", "warning");
        event.preventDefault();
        return;
    }

    event.preventDefault(); // Impede a colagem padrão

    const pastedText = (event.clipboardData || window.clipboardData).getData("text/plain");
    const lines = pastedText.trim().split(/\r?\n|\r/).filter(line => line.length > 0);

    if (lines.length === 0) {
        Swal.fire("⚠️ Nada para colar", "Nenhum dado válido encontrado na área de transferência.", "info");
        return;
    }

    Swal.fire({
        title: "Colando dados...",
        text: `Serão colados ${lines.length} itens a partir da linha ${rowIndex + 1}.`,
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    for (let i = 0; i < lines.length; i++) {
        let currentLineIndex = rowIndex + i;
        let targetRow = tabela.rows[currentLineIndex];

        if (!targetRow) {
            const newRow = criarLinhaVazia();
            tabela.appendChild(newRow);
            targetRow = newRow;
        }

        const inputToUpdate = targetRow.querySelectorAll("td")[columnIndex]?.querySelector("input");
        if (inputToUpdate) {
            inputToUpdate.value = lines[i].toUpperCase();
            if (columnIndex === 2) { // Se a coluna de nível foi colada, aplica indentação
                aplicarIndentacao(targetRow);
            }
        }
    }

    atualizarSequencias();
    verificarDuplicatas();
    Swal.fire("✅ Colagem concluída", `${lines.length} itens colados com sucesso!`, "success");
}

// Adiciona o event listener para o paste em toda a tabela
document.getElementById("listaTabela").addEventListener("paste", handlePasteMultipleLines);


// Função para determinar a cor do texto para contraste
function getContrastYIQ(hexcolor) {
    var r = parseInt(hexcolor.substr(1, 2), 16);
    var g = parseInt(hexcolor.substr(3, 2), 16);
    var b = parseInt(hexcolor.substr(5, 2), 16);
    var yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
    return (yiq >= 128) ? 'black' : 'white';
}

document.addEventListener("DOMContentLoaded", () => {
    if (tabela.rows.length === 0) {
        criar10Linhas();
        atualizarSequencias();
        verificarDuplicatas();
        Swal.fire("🎉 Bem-vindo!", "A lista foi inicializada com 10 linhas para você começar.", "info");
    }

    // Aplica as classes iniciais de visibilidade das colunas
    const tabelaElement = document.getElementById("listaTabela");
    if (!seqAtivo) {
        tabelaElement.classList.add("seq-col-hidden");
    }
    if (!nivelColVisivel) {
        tabelaElement.classList.add("nivel-col-hidden");
    }

    // Inicializa o estado da régua
    if (hoverEffectAtivo) {
        tabelaElement.classList.add("hover-effect");
    } else {
        tabelaElement.classList.add("no-hover-effect");
    }
});