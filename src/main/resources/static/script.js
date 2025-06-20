let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true;
let corSelecionada = "";
let demarcarLinha = false;
let removerDemarcacao = false;


const unidades = ["un", "cj", "kg", "mm", "m"];
const tiposEstrutura = ["Manufatura", "Comprado", ""];
const fatorSucata = ["0", "15", ""];
const linhaValores = Array.from({ length: 80 }, (_, i) => String((i + 1) * 10));
const alternativas = ["*", ""];
const siteValores = ["1", ""];
const niveis = Array.from({ length: 10 }, (_, i) => (i + 1).toString());

function inputCell(type, readOnly = false, value = "", isPasteTarget = false) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = type;
    input.readOnly = readOnly;
    input.value = (value || "").toUpperCase();
    input.addEventListener("input", (e) => {
        e.target.value = e.target.value.toUpperCase();
        verificarDuplicatas();
    });

    // Evento de clique para pintura de célula (com nova lógica)
    input.addEventListener("click", (e) => {
        if (!demarcarLinha) { // Se não for para demarcar a linha inteira (ou seja, demarcar célula)
            if (removerDemarcacao) {
                e.target.style.backgroundColor = ""; // Sempre remove
            } else if (corSelecionada) {
                // Toggle de pintura na célula
                if (e.target.style.backgroundColor === corSelecionada) {
                    e.target.style.backgroundColor = ""; // Remove a cor
                } else {
                    e.target.style.backgroundColor = corSelecionada; // Aplica a cor
                }
                // corSelecionada NÃO É MAIS LIMPA AQUI para permitir demarcar várias
            }
        }
    });

    if (isPasteTarget) {
        input.addEventListener("paste", handlePasteMultipleLines);
    }

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
        if (!demarcarLinha) { // Se não for para demarcar a linha inteira (ou seja, demarcar célula)
            if (removerDemarcacao) {
                e.target.style.backgroundColor = ""; // Sempre remove
            } else if (corSelecionada) {
                // Toggle de pintura na célula
                if (e.target.style.backgroundColor === corSelecionada) {
                    e.target.style.backgroundColor = ""; // Remove a cor
                } else {
                    e.target.style.backgroundColor = corSelecionada; // Aplica a cor
                }
                // corSelecionada NÃO É MAIS LIMPA AQUI para permitir demarcar várias
            }
        }
    });
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

    const indentTd = document.createElement("td");
    const btnMais = document.createElement("button");
    const btnMenos = document.createElement("button");
    btnMais.textContent = "➕";
    btnMenos.textContent = "➖";
    btnMais.addEventListener("click", () => ajustarNivel(row, 1));
    btnMenos.addEventListener("click", () => ajustarNivel(row, -1));
    indentTd.appendChild(btnMenos);
    indentTd.appendChild(btnMais);
    row.appendChild(indentTd);

    const seqTd = document.createElement("td");
    seqTd.classList.add("seq-col");
    row.appendChild(seqTd);

    const nivelTd = document.createElement("td");
    const nivelSelect = document.createElement("select");
    nivelSelect.classList.add("nivel-select");
    niveis.forEach(n => {
        const opt = document.createElement("option");
        opt.value = n;
        opt.textContent = n;
        nivelSelect.appendChild(opt);
    });
    nivelSelect.value = v.NIVEL || "1";
    nivelSelect.addEventListener("change", () => aplicarIndentacao(row));
    nivelTd.appendChild(nivelSelect);
    row.appendChild(nivelTd);

    // Adiciona evento de clique para a linha inteira para pintar (com nova lógica)
    row.addEventListener("click", (e) => {
        if (demarcarLinha && !e.target.tagName.match(/INPUT|SELECT|BUTTON/)) { // Apenas se "Demarcar linha" estiver ativada
            if (removerDemarcacao) {
                row.style.backgroundColor = ""; // Sempre remove
            } else if (corSelecionada) {
                // Toggle de pintura na linha
                if (row.style.backgroundColor === corSelecionada) {
                    row.style.backgroundColor = ""; // Remove a cor
                } else {
                    row.style.backgroundColor = corSelecionada; // Aplica a cor
                }
                // corSelecionada NÃO É MAIS LIMPA AQUI para permitir demarcar várias
            }
        }
    });

    row.appendChild(selectCell(siteValores, v.SITE || "1"));
    row.appendChild(selectCell(alternativas, v.ALTERNATIVA || "*"));
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || "", true));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));
    row.appendChild(selectCell(linhaValores, v.LINHA || "10"));
    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || "", true));
    // QTDE_MONTAGEM agora é tipo "text" e paste target
    row.appendChild(inputCell("text", false, v.QTDE_MONTAGEM || "0", true)); // Alterado para type="text" e isPasteTarget = true
    row.appendChild(selectCell(unidades, v.UNIDADE_MEDIDA || "un"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    aplicarIndentacao(row);

    if (!seqAtivo) {
        row.querySelector(".seq-col").style.display = "none";
    }

    return row;
}

function criarLinhaVazia() {
    return criarLinha({});
}

function atualizarSequencias() {
    const linhas = tabela.querySelectorAll("tr");
    linhas.forEach((row, index) => {
        const seqTd = row.querySelector("td.seq-col");
        if (seqTd) {
            seqTd.textContent = (index + 1) * 10;
        }
    });
}

function ajustarNivel(row, delta) {
    const select = row.querySelector(".nivel-select");
    if (!select) return;
    let nivelAtual = parseInt(select.value);
    nivelAtual = Math.min(10, Math.max(1, nivelAtual + delta));
    select.value = nivelAtual;
    aplicarIndentacao(row);
}

function aplicarIndentacao(row) {
    for (let i = 1; i <= 10; i++) {
        row.classList.remove(`nivel-${i}`);
    }
    const nivel = row.querySelector(".nivel-select")?.value;
    if (nivel) {
        row.classList.add(`nivel-${nivel}`);
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
        SITE: cells[4]?.querySelector("select")?.value || "",
        ALTERNATIVA: cells[5]?.querySelector("select")?.value || "",
        CODIGO_MATERIAL: cells[6]?.querySelector("input")?.value.trim().toUpperCase() || "",
        NIVEL: cells[3]?.querySelector("select")?.value || "1",
        TIPO_ESTRUTURA: cells[7]?.querySelector("select")?.value || "",
        LINHA: cells[8]?.querySelector("select")?.value || "",
        ITEM_COMPONENTE: cells[9]?.querySelector("input")?.value.trim().toUpperCase() || "",
        QTDE_MONTAGEM: cells[10]?.querySelector("input")?.value.trim() || "",
        UNIDADE_MEDIDA: cells[11]?.querySelector("select")?.value || "",
        FATOR_SUCATA: cells[12]?.querySelector("select")?.value || ""
    };
}

function preencherLinha(row, data) {
    const cells = row.querySelectorAll("td");
    cells[3].querySelector("select").value = data.NIVEL || "1";
    aplicarIndentacao(row);
    cells[4].querySelector("select").value = data.SITE || "1";
    cells[5].querySelector("select").value = data.ALTERNATIVA || "*";
    cells[6].querySelector("input").value = (data.CODIGO_MATERIAL || "").toUpperCase();
    cells[7].querySelector("select").value = data.TIPO_ESTRUTURA || "Manufatura";
    cells[8].querySelector("select").value = data.LINHA || "10";
    cells[9].querySelector("input").value = (data.ITEM_COMPONENTE || "").toUpperCase();
    cells[10].querySelector("input").value = data.QTDE_MONTAGEM || "0";
    cells[11].querySelector("select").value = data.UNIDADE_MEDIDA || "un";
    cells[12].querySelector("select").value = data.FATOR_SUCATA || "0";
}

function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    const hashes = new Map();
    const conjuntos = new Map();

    linhas.forEach(row => {
        // Apenas reseta cores de fundo de duplicata se não for uma cor de demarcação manual
        const currentColor = row.style.backgroundColor;
        const demarcationColors = ["rgb(208, 235, 255)", "rgb(208, 240, 192)", "rgb(173, 216, 230)", "rgb(255, 204, 204)"];
        if (currentColor === "rgb(240, 230, 255)" || currentColor === "rgb(217, 194, 255)") { // Cores de duplicata
            row.style.backgroundColor = "";
        }
    });

    linhas.forEach((tr) => {
        const data = getLinhaData(tr);

        if (data.CODIGO_MATERIAL === "" && data.ITEM_COMPONENTE === "") return;

        const hash = `${data.SITE}|${data.ALTERNATIVA}|${data.CODIGO_MATERIAL}|${data.NIVEL}|${data.TIPO_ESTRUTURA}|${data.LINHA}|${data.ITEM_COMPONENTE}|${data.QTDE_MONTAGEM}|${data.UNIDADE_MEDIDA}|${data.FATOR_SUCATA}`;
        const groupHash = `${data.CODIGO_MATERIAL}|${data.ITEM_COMPONENTE}`;

        if (!hashes.has(hash)) hashes.set(hash, []);
        hashes.get(hash).push(tr);

        if (!conjuntos.has(groupHash)) conjuntos.set(groupHash, []);
        conjuntos.get(groupHash).push(tr);
    });

    for (const [hash, rows] of hashes) {
        if (rows.length > 1) {
            rows.forEach(row => {
                const currentColor = row.style.backgroundColor;
                const demarcationColors = ["rgb(208, 235, 255)", "rgb(208, 240, 192)", "rgb(173, 216, 230)", "rgb(255, 204, 204)"];
                if (!demarcationColors.includes(currentColor)) { // Não sobrescreve cores de demarcação manual
                     row.style.backgroundColor = "#f0e6ff"; // Roxo claro
                }
            });
        }
    }

    for (const [groupHash, rows] of conjuntos) {
        if (rows.length > 1) {
            rows.forEach(row => {
                const currentColor = row.style.backgroundColor;
                const demarcationColors = ["rgb(208, 235, 255)", "rgb(208, 240, 192)", "rgb(173, 216, 230)", "rgb(255, 204, 204)"];
                if (currentColor !== "rgb(240, 230, 255)" && !demarcationColors.includes(currentColor)) {
                    row.style.backgroundColor = "#d9c2ff"; // Roxo médio
                }
            });
        }
    }
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
    const copiarLinhaCheckbox = document.getElementById("copiarLinhaCheckbox").checked;
    const copiarConjuntoCheckbox = document.getElementById("copiarConjuntoCheckbox").checked;
    const selecionados = document.querySelectorAll(".linha-selecao:checked");

    if (selecionados.length === 0) {
        Swal.fire("⚠️ Nada selecionado", "Marque uma linha.", "info");
        return;
    }

    cacheCopiado = [];

    if (copiarLinhaCheckbox) {
        cacheCopiado = Array.from(selecionados).map(cb => getLinhaData(cb.closest("tr")));
    } else if (copiarConjuntoCheckbox) {
        if (selecionados.length !== 1) {
            Swal.fire("⚠️ Selecione uma única linha", "Para copiar o conjunto, selecione apenas uma linha.", "info");
            return;
        }
        const linhaBaseData = getLinhaData(selecionados[0].closest("tr"));
        if (!linhaBaseData.CODIGO_MATERIAL && !linhaBaseData.ITEM_COMPONENTE) {
            Swal.fire("⚠️ Linha vazia", "Para copiar o conjunto, o Código Material ou Item Componente da linha selecionada não pode estar vazio.", "info");
            return;
        }

        const todasAsLinhas = Array.from(tabela.rows);
        cacheCopiado = todasAsLinhas.filter(tr => {
            const data = getLinhaData(tr);
            return (data.CODIGO_MATERIAL === linhaBaseData.CODIGO_MATERIAL && data.ITEM_COMPONENTE === linhaBaseData.ITEM_COMPONENTE);
        }).map(tr => getLinhaData(tr));

    } else {
        Swal.fire("⚠️ Selecione uma opção de cópia", "Escolha 'Copiar Linha' ou 'Copiar Conjunto'.", "info");
        return;
    }

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
                SITE: rowData.SITE || rowData.Site || "1",
                ALTERNATIVA: rowData.ALTERNATIVA || rowData.Alternativa || "*",
                CODIGO_MATERIAL: rowData["CÓDIGO_MATERIAL"] || rowData["CODIGO_MATERIAL"] || rowData["CODIGO MATERIAL"] || rowData.Codigo_Material || rowData.CodigoMaterial || "",
                NIVEL: String(rowData.NIVEL || rowData.Nivel || "1"),
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


document.getElementById("toggleSeqBtn").addEventListener("click", () => {
    const seqHeader = document.querySelector("#listaTabela th.seq-col");
    const seqCells = document.querySelectorAll("#listaTabela td.seq-col");
    const icon = document.querySelector("#toggleSeqBtn .bi");

    if (seqAtivo) {
        seqHeader.style.display = "none";
        seqCells.forEach(cell => cell.style.display = "none");
        icon.classList.remove("bi-eye");
        icon.classList.add("bi-eye-slash");
        seqAtivo = false;
    } else {
        seqHeader.style.display = "";
        seqCells.forEach(cell => cell.style.display = "");
        icon.classList.remove("bi-eye-slash");
        icon.classList.add("bi-eye");
        seqAtivo = true;
    }
});

// Lógica para pintar
document.querySelectorAll(".paint-btn").forEach(button => {
    button.addEventListener("click", function() {
        corSelecionada = this.dataset.color;
        Swal.fire({
            title: "Cor selecionada!",
            text: `Clique na ${demarcarLinha ? 'linha' : 'célula'} que deseja pintar com ${corSelecionada}. Clique novamente para remover a cor.`,
            icon: "info",
            timer: 3000,
            showConfirmButton: false
        });
    });
});

// Event listener para a checkbox "Demarcar linha"
document.getElementById("demarcarLinhaCheckbox").addEventListener("change", function() {
    demarcarLinha = this.checked;
    if (demarcarLinha) {
        Swal.fire("Demarcar linha ativado!", "Clique em um botão de cor e depois na linha para demarcar/desdemarcar.", "info");
    } else {
        Swal.fire("Demarcar célula ativado!", "Clique em um botão de cor e depois na célula para demarcar/desdemarcar.", "info");
    }
});

// Event listener para a nova checkbox "Remover demarcação"
document.getElementById("removerDemarcacaoCheckbox").addEventListener("change", function() {
    removerDemarcacao = this.checked;
    if (removerDemarcacao) {
        Swal.fire("Modo de remoção de demarcação ativado!", "Agora, cliques apagarão as demarcações existentes.", "warning");
        corSelecionada = ""; // Limpa a cor selecionada para evitar aplicação acidental
    } else {
        Swal.fire("Modo de remoção de demarcação desativado!", "Pode voltar a demarcar.", "info");
    }
});

document.getElementById("clearPaintBtn").addEventListener("click", () => {
    corSelecionada = "";
    document.getElementById("removerDemarcacaoCheckbox").checked = false; // Desmarca remover demarcação
    removerDemarcacao = false;
    Swal.fire("Seleção de cor limpa!", "Agora o clique não aplicará cores.", "info", 1500);
});


document.getElementById("toggleAllCheckboxes").addEventListener("change", function() {
    const checkboxes = document.querySelectorAll(".linha-selecao");
    checkboxes.forEach(cb => {
        cb.checked = this.checked;
    });
});


async function handlePasteMultipleLines(event) {
    if (corSelecionada || demarcarLinha || removerDemarcacao) { // Verifica todas as condições de pintura
        Swal.fire("Modo de Demarcação Ativo", "Desative o modo de demarcação ou limpe a cor selecionada para colar.", "warning");
        event.preventDefault();
        return;
    }

    const targetCell = event.target;
    const td = targetCell.closest("td");
    const tr = targetCell.closest("tr");
    const rowIndex = Array.from(tabela.rows).indexOf(tr);

    let columnIndex = -1;
    // Identifica qual coluna está sendo colada (CODIGO_MATERIAL ou ITEM_COMPONENTE ou QTDE_MONTAGEM)
    if (td === tr.querySelectorAll("td")[6]) {
        columnIndex = 6;
    } else if (td === tr.querySelectorAll("td")[9]) {
        columnIndex = 9;
    } else if (td === tr.querySelectorAll("td")[10]) { // Adicionado para QTDE_MONTAGEM
        columnIndex = 10;
    } else {
        return;
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
        atualizarSequencias(); // Garante que a sequência esteja correta na inicialização
        verificarDuplicatas();
        Swal.fire("🎉 Bem-vindo!", "A lista foi inicializada com 10 linhas para você começar.", "info");
    }
});