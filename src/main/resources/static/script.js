let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true; // SEQ estará visível por padrão
let nivelColVisivel = true; // Nível estará visível por padrão
let corSelecionada = ""; // Cor hex (#RRGGBB) para aplicar
let demarcarLinha = false; // true: clica na linha, false: clica na célula
let removerDemarcacao = false; // true: cliques removem cores, false: cliques aplicam cores
let ignorarDuplicatas = false; // Nova variável para ignorar destaque de duplicatas

// Definição das cores para os níveis
const nivelColors = [
    "#ffe0b2", // Nível 1 - Pastel
    "#bfefbb", // Nível 2 - Pastel
    "#c0d9ef", // Nível 3 - Pastel
    "#e6b2e6", // Nível 4 - Pastel
    "#ffe0e0", // Nível 5 - Pastel
    "#a7d9b5", // Nível 6 - Pastel
    "#d5e6a7", // Nível 7 - Pastel
    "#b2e0e0", // Nível 8 - Pastel
    "#e0b2a7", // Nível 9 - Pastel
    "#a7b5d9"  // Nível 10 - Pastel
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

    // SEQ agora é a primeira coluna de dados (índice 1 no array de células)
    const seqTd = document.createElement("td");
    seqTd.classList.add("seq-col");
    row.appendChild(seqTd);

    // Nível agora é a segunda coluna de dados (índice 2 no array de células)
    const nivelCell = inputCell("text", false, v.NIVEL || "", true, "nivel-col");
    row.appendChild(nivelCell);

    // Event listener para demarcação da LINHA
    row.addEventListener("click", (e) => {
        // Impede a demarcação se o clique for em input/select (para permitir digitação)
        if (e.target.tagName.match(/INPUT|SELECT|BUTTON/)) {
            return;
        }

        if (demarcarLinha) { // Modo de demarcação de linha
            if (removerDemarcacao) {
                // Se removerDemarcacao estiver ON, sempre remove a cor
                row.style.backgroundColor = "";
            } else if (corSelecionada) {
                // Se corSelecionada estiver definida, aplica/remove essa cor
                if (row.style.backgroundColor === corSelecionada) {
                    row.style.backgroundColor = ""; // Remove se já tiver a cor
                } else {
                    row.style.backgroundColor = corSelecionada; // Aplica a cor
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

    // Esconde SEQ se estiver desativado
    if (!seqAtivo) {
        seqTd.style.display = "none";
    }
    // Esconde NÍVEL se estiver desativado
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
        // A coluna SEQ é agora a segunda TD na linha (índice 1)
        const seqTd = row.querySelectorAll("td")[1];
        if (seqTd) {
            seqTd.textContent = (index + 1) * 10;
        }
    });
}

function aplicarIndentacao(row) {
    for (let i = 1; i <= 10; i++) {
        row.classList.remove(`nivel-${i}`);
    }
    // A coluna NÍVEL é agora a terceira TD na linha (índice 2)
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
        // Ajuste dos índices das células devido à nova ordem:
        // cells[0] é a checkbox
        // cells[1] é SEQ
        // cells[2] é NÍVEL
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
    // Ajuste dos índices das células:
    // cells[0] é a checkbox
    // cells[1] é SEQ
    // cells[2] é NÍVEL
    cells[2].querySelector("input").value = data.NIVEL || "";
    aplicarIndentacao(row); // Reaplicar indentação ao preencher
    cells[3].querySelector("select").value = data.SITE || "1";
    cells[4].querySelector("select").value = data.ALTERNATIVA || "*";
    cells[5].querySelector("input").value = (data.CODIGO_MATERIAL || "").toUpperCase();
    cells[6].querySelector("select").value = data.TIPO_ESTRUTURA || "Manufatura";
    cells[7].querySelector("select").value = data.LINHA || "10";
    cells[8].querySelector("input").value = (data.ITEM_COMPONENTE || "").toUpperCase();
    cells[9].querySelector("input").value = (data.QTDE_MONTAGEM || "0").replace(",", "."); // Garante ponto decimal
    cells[10].querySelector("select").value = data.UNIDADE_MEDIDA || "un";
    cells[11].querySelector("select").value = data.FATOR_SUCATA || "0";
}

function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    const hashes = new Map();
    // Cores de demarcação padrão, para não serem confundidas com a demarcação de duplicatas
    const defaultDemarcationColors = nivelColors.concat([
        "#ffc107", // Amarelo Atenção
        "#4CAF50", // Verde OK
        "#f44336"  // Vermelho Crítico
    ]).map(c => rgbToHex(c)); // Converte para hex para comparação consistente

    // Reseta a cor de duplicata exata, se houver
    linhas.forEach(row => {
        const currentColor = row.style.backgroundColor;
        // Converte a cor atual da linha para HEX para comparação.
        // Se a cor atual for "#f0e6ff" (roxo claro), a gente remove.
        if (rgbToHex(currentColor) === "#F0E6FF") {
            row.style.backgroundColor = "";
        }
    });

    if (ignorarDuplicatas) {
        return;
    }

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
                const currentColor = row.style.backgroundColor;
                // Aplica a cor de duplicata SOMENTE se a linha não estiver com uma cor de demarcação manual
                if (!defaultDemarcationColors.includes(rgbToHex(currentColor))) {
                    row.style.backgroundColor = "#f0e6ff"; // Roxo claro para duplicatas exatas
                }
            });
        }
    }
}

// Helper para converter RGB para HEX (para consistência na comparação de cores)
function rgbToHex(rgb) {
    if (!rgb || rgb.indexOf('rgb') === -1) {
        return rgb ? rgb.toUpperCase() : ""; // Retorna a string original se não for RGB (pode ser HEX direto ou vazio)
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

    // Definir largura da coluna NÍVEL (coluna C no Excel, índice 2)
    ws['!cols'] = ws['!cols'] || [];
    ws['!cols'][2] = { wch: 5 }; // Define a largura para a coluna do NÍVEL (coluna C)
    ws['!cols'][1] = { wch: 5 }; // Define a largura para a coluna SEQ (coluna B)


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

// Lógica para os botões de cor de Nível
const nivelColorButtonsContainer = document.getElementById("nivelColorButtons");
nivelColors.forEach((color, index) => {
    const button = document.createElement("button");
    button.className = `paint-btn btn-nivel-${index + 1}`;
    button.dataset.color = color;
    button.textContent = index + 1; // Número do nível
    button.addEventListener("click", function() {
        corSelecionada = this.dataset.color;
        document.getElementById("removerDemarcacaoCheckbox").checked = false; // Desativa remoção
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

// Lógica para os botões de cor de Atenção
document.querySelectorAll("#attentionColorButtons .paint-btn").forEach(button => {
    button.addEventListener("click", function() {
        corSelecionada = this.dataset.color;
        document.getElementById("removerDemarcacaoCheckbox").checked = false; // Desativa remoção
        removerDemarcacao = false;
        const label = this.nextElementSibling ? this.nextElementSibling.textContent : "uma cor";
        Swal.fire({
            title: `Cor "${label}" selecionada!`,
            text: `Clique na ${demarcarLinha ? 'linha' : 'célula'} que deseja pintar.`,
            icon: "info",
            timer: 2000,
            showConfirmButton: false
        });
    });
});


// Event listener para a checkbox "Demarcar linha"
document.getElementById("demarcarLinhaCheckbox").addEventListener("change", function() {
    demarcarLinha = this.checked;
    if (demarcarLinha) {
        Swal.fire("Modo 'Demarcar linha' ativado!", "Agora, cliques aplicarão a cor selecionada à linha inteira.", "info");
    } else {
        Swal.fire("Modo 'Demarcar linha' desativado!", "Cliques demarcarão células individuais.", "info");
    }
});

// Event listener para a nova checkbox "Remover demarcação"
document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", function() {
    removerDemarcacao = this.checked;
    if (removerDemarcacao) {
        corSelecionada = ""; // Limpa a cor selecionada para garantir que só remova
        Swal.fire("Modo 'Remover demarcação' ativado!", "Agora, cliques removerão as demarcações existentes. Selecione um botão de cor para sair deste modo.", "warning");
    } else {
        Swal.fire("Modo 'Remover demarcação' desativado!", "Pode voltar a demarcar.", "info");
    }
});

document.getElementById("clearPaintBtn").addEventListener("click", () => {
    corSelecionada = "";
    document.getElementById("removerDemarcacaoCheckbox").checked = false; // Desmarca remover demarcação
    demarcarLinha = false; // Garante que volta para demarcação de célula
    document.getElementById("demarcarLinhaCheckbox").checked = false;
    removerDemarcacao = false;
    Swal.fire("Seleção de cor limpa!", "Agora o clique não aplicará cores.", "info", 1500);
});

// Lógica para o checkbox "Ignorar destaque de duplicatas"
document.getElementById("ignorarDuplicatasCheckbox").addEventListener("change", function() {
    ignorarDuplicatas = this.checked;
    verificarDuplicatas(); // Re-verifica duplicatas para aplicar/remover o destaque
    if (ignorarDuplicatas) {
        Swal.fire("Destaque de duplicatas desativado!", "As linhas duplicadas não serão mais destacadas com roxo claro.", "info", 2000);
    } else {
        Swal.fire("Destaque de duplicatas ativado!", "Linhas duplicadas serão destacadas com roxo claro.", "info", 2000);
    }
});

// Lógica para Marcar/Desmarcar Todos com SweetAlert para apagar demarcações
document.getElementById("toggleAllCheckboxes").addEventListener("change", function() {
    const checkboxes = document.querySelectorAll(".linha-selecao");
    const headerCheckbox = document.getElementById("toggleAllCheckboxesHeader");

    // Sincroniza o checkbox do cabeçalho
    headerCheckbox.checked = this.checked;

    if (this.checked) {
        checkboxes.forEach(cb => {
            cb.checked = true;
        });
    } else {
        Swal.fire({
            title: "Apagar todas as demarcações?",
            text: "Deseja remover todas as cores aplicadas à tabela?",
            icon: "warning",
            showCancelButton: true,
            confirmButtonColor: "#3085d6",
            cancelButtonColor: "#d33",
            confirmButtonText: "Sim, apagar!",
            cancelButtonText: "Não, cancelar"
        }).then((result) => {
            if (result.isConfirmed) {
                // Remove todas as demarcações
                const allCells = tabela.querySelectorAll("td");
                const allRows = tabela.querySelectorAll("tr");
                allCells.forEach(cell => cell.style.backgroundColor = "");
                allRows.forEach(row => row.style.backgroundColor = "");

                // Desmarca todas as checkboxes
                checkboxes.forEach(cb => cb.checked = false);
                headerCheckbox.checked = false; // Garante que o cabeçalho também seja desmarcado
                Swal.fire("Demarcações removidas!", "", "success", 1500);
            } else {
                // Se o usuário cancelar, re-seleciona "Marcar/Desmarcar Todos"
                this.checked = true;
                headerCheckbox.checked = true;
                checkboxes.forEach(cb => cb.checked = true); // Mantém todas as checkboxes marcadas
            }
        });
    }
});

// Sincroniza o checkbox do cabeçalho com o checkbox de "Marcar/Desmarcar Todos"
document.getElementById("toggleAllCheckboxesHeader").addEventListener("change", function() {
    const mainToggleCheckbox = document.getElementById("toggleAllCheckboxes");
    mainToggleCheckbox.checked = this.checked;

    // Dispara o evento change para que a lógica principal seja executada
    mainToggleCheckbox.dispatchEvent(new Event('change'));
});

// Lógica para exibir/ocultar a coluna SEQ
document.getElementById("toggleSeqBtn").addEventListener("click", () => {
    const seqHeader = document.querySelector("th.seq-col");
    const seqCells = document.querySelectorAll("td.seq-col");

    seqAtivo = !seqAtivo;

    seqHeader.style.display = seqAtivo ? "" : "none";
    seqCells.forEach(cell => {
        cell.style.display = seqAtivo ? "" : "none";
    });

    if (seqAtivo) {
        Swal.fire("SEQ Visível!", "A coluna SEQ está visível.", "info", 1500);
    } else {
        Swal.fire("SEQ Oculta!", "A coluna SEQ está oculta.", "info", 1500);
    }
});

// Lógica para exibir/ocultar a coluna NÍVEL
document.getElementById("toggleNivelColBtn").addEventListener("click", () => {
    const nivelHeader = document.querySelector("th.nivel-col");
    const nivelCells = document.querySelectorAll("td.nivel-col");

    nivelColVisivel = !nivelColVisivel;

    nivelHeader.style.display = nivelColVisivel ? "" : "none";
    nivelCells.forEach(cell => {
        cell.style.display = nivelColVisivel ? "" : "none";
    });

    if (nivelColVisivel) {
        Swal.fire("NÍVEL Visível!", "A coluna NÍVEL está visível.", "info", 1500);
    } else {
        Swal.fire("NÍVEL Oculta!", "A coluna NÍVEL está oculta.", "info", 1500);
    }
});


async function handlePasteMultipleLines(event) {
    if (corSelecionada || demarcarLinha || removerDemarcacao) {
        Swal.fire("Modo de Demarcação Ativo", "Desative o modo de demarcação ou limpe a cor selecionada para colar.", "warning");
        event.preventDefault();
        return;
    }

    const targetCell = event.target;
    const td = targetCell.closest("td");
    const tr = targetCell.closest("tr");
    const rowIndex = Array.from(tabela.rows).indexOf(tr);

    let columnIndex = -1;
    // Ajustar os índices das células para o NÍVEL, CÓDIGO_MATERIAL, ITEM_COMPONENTE e QTDE_MONTAGEM
    // Lembre-se que o índice 0 é a checkbox de seleção de linha
    // O índice 1 é SEQ
    // O índice 2 é NÍVEL
    const allCellsInRow = tr.querySelectorAll("td");
    if (targetCell === allCellsInRow[2]?.querySelector("input")) { // NÍVEL (index 2)
        columnIndex = 2;
    } else if (targetCell === allCellsInRow[5]?.querySelector("input")) { // CÓDIGO_MATERIAL (index 5)
        columnIndex = 5;
    } else if (targetCell === allCellsInRow[8]?.querySelector("input")) { // ITEM_COMPONENTE (index 8)
        columnIndex = 8;
    } else if (targetCell === allCellsInRow[9]?.querySelector("input")) { // QTDE_MONTAGEM (index 9)
        columnIndex = 9;
    } else {
        return; // Não é uma coluna para colar múltiplos itens
    }

    event.preventDefault();

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
            if (columnIndex === 2) { // Se a coluna for NÍVEL, aplica a indentação imediatamente
                aplicarIndentacao(targetRow);
            }
        }
    }

    atualizarSequencias();
    verificarDuplicatas();
    Swal.fire("✅ Colagem concluída", `${lines.length} itens colados com sucesso!`, "success");
}


// Ao carregar a página, cria 10 linhas vazias para iniciar e exibe a mensagem
document.addEventListener("DOMContentLoaded", () => {
    if (tabela.rows.length === 0) {
        criar10Linhas();
        atualizarSequencias();
        verificarDuplicatas();
        Swal.fire("🎉 Bem-vindo!", "A lista foi inicializada com 10 linhas para você começar.", "info");
    }

    // Inicializa o estado de visibilidade das colunas ao carregar
    const seqHeader = document.querySelector("th.seq-col");
    const seqCells = document.querySelectorAll("td.seq-col");
    if (!seqAtivo) {
        seqHeader.style.display = "none";
        seqCells.forEach(cell => cell.style.display = "none");
    }

    const nivelHeader = document.querySelector("th.nivel-col");
    const nivelCells = document.querySelectorAll("td.nivel-col");
    if (!nivelColVisivel) {
        nivelHeader.style.display = "none";
        nivelCells.forEach(cell => cell.style.display = "none");
    }
});