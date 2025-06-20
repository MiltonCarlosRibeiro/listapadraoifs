let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true;
let corSelecionada = "";
let pintarLinhaInteira = false; // Variável renomeada para clareza

const unidades = ["un", "cj", "kg", "mm", "m"];
const tiposEstrutura = ["Manufatura", "Comprado", ""];
const fatorSucata = ["0", "15", ""];
const linhaValores = Array.from({ length: 80 }, (_, i) => String((i + 1) * 10));
const alternativas = ["*", ""];
const siteValores = ["1", ""];
const niveis = Array.from({ length: 10 }, (_, i) => (i + 1).toString());

function inputCell(type, readOnly = false, value = "") {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = type;
    input.readOnly = readOnly;
    input.value = (value || "").toUpperCase(); // Garante string vazia para undefined/null
    input.addEventListener("input", (e) => {
        e.target.value = e.target.value.toUpperCase();
        verificarDuplicatas();
    });
    // Adiciona evento para aplicar cor se uma cor estiver selecionada
    input.addEventListener("click", (e) => {
        if (corSelecionada && !pintarLinhaInteira) {
            e.target.style.backgroundColor = corSelecionada;
            corSelecionada = ""; // Limpa a cor após aplicar
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
    // Adiciona evento para aplicar cor se uma cor estiver selecionada
    select.addEventListener("click", (e) => {
        if (corSelecionada && !pintarLinhaInteira) {
            e.target.style.backgroundColor = corSelecionada;
            corSelecionada = ""; // Limpa a cor após aplicar
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
    seqTd.classList.add("seq-col"); // Adicionado classe para ser ocultado/exibido
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

    // Adiciona evento de clique para a linha inteira para pintar
    row.addEventListener("click", (e) => {
        if (corSelecionada && pintarLinhaInteira) {
            row.style.backgroundColor = corSelecionada;
            corSelecionada = ""; // Limpa a cor após aplicar
        }
    });

    row.appendChild(selectCell(siteValores, v.SITE || "1"));
    row.appendChild(selectCell(alternativas, v.ALTERNATIVA || "*"));
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || ""));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));
    row.appendChild(selectCell(linhaValores, v.LINHA || "10"));
    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || ""));
    row.appendChild(inputCell("number", false, v.QTDE_MONTAGEM || "0"));
    row.appendChild(selectCell(unidades, v.UNIDADE_MEDIDA || "un"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    aplicarIndentacao(row);
    tabela.appendChild(row);
    atualizarSequencias();

    // Se SEQ estiver oculto, esconde a nova linha também
    if (!seqAtivo) {
        row.querySelector(".seq-col").style.display = "none";
    }
}

function criarLinhaVazia() {
    const dummy = document.createElement("tbody");
    criarLinha();
    dummy.appendChild(tabela.lastChild);
    return dummy.removeChild(dummy.firstChild);
}

function atualizarSequencias() {
    const linhas = tabela.querySelectorAll("tr");
    linhas.forEach((row, index) => {
        const seqTd = row.querySelector("td.seq-col"); // Usar a classe correta
        if (seqTd) {
            seqTd.textContent = (index + 1) * 10;
        }
    });
}

function ajustarNivel(row, delta) {
    const select = row.querySelector(".nivel-select");
    if (!select) return; // Garante que o elemento existe
    let nivelAtual = parseInt(select.value);
    nivelAtual = Math.min(10, Math.max(1, nivelAtual + delta));
    select.value = nivelAtual;
    aplicarIndentacao(row);
}

function aplicarIndentacao(row) {
    for (let i = 1; i <= 10; i++) {
        row.classList.remove(`nivel-${i}`);
    }
    const nivel = row.querySelector(".nivel-select")?.value; // Usar optional chaining
    if (nivel) {
        row.classList.add(`nivel-${nivel}`);
    }
}

function criar10Linhas() {
    for (let i = 0; i < 10; i++) criarLinha();
}

function getLinhaData(tr) {
    const cells = tr.querySelectorAll("td");
    // Adicionado verificação de existência para cada célula/elemento
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
    // Usar optional chaining e garantir valor padrão
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
    const hashes = new Map(); // Para duplicatas exatas
    const conjuntos = new Map(); // Para duplicatas de conjunto (código material + item componente)

    // Resetar cores
    linhas.forEach(row => {
        row.style.backgroundColor = "";
        Array.from(row.querySelectorAll("td input, td select")).forEach(el => el.style.backgroundColor = "");
    });

    linhas.forEach((tr) => {
        const data = getLinhaData(tr);

        // Apenas considera linhas com pelo menos um dos campos chave preenchido
        if (data.CODIGO_MATERIAL === "" && data.ITEM_COMPONENTE === "") return;

        // Hash para duplicatas exatas (todos os campos chave)
        const hash = `${data.SITE}|${data.ALTERNATIVA}|${data.CODIGO_MATERIAL}|${data.NIVEL}|${data.TIPO_ESTRUTURA}|${data.LINHA}|${data.ITEM_COMPONENTE}|${data.QTDE_MONTAGEM}|${data.UNIDADE_MEDIDA}|${data.FATOR_SUCATA}`;
        // Hash para duplicatas de conjunto (código material e item componente)
        const groupHash = `${data.CODIGO_MATERIAL}|${data.ITEM_COMPONENTE}`;

        if (!hashes.has(hash)) hashes.set(hash, []);
        hashes.get(hash).push(tr);

        if (!conjuntos.has(groupHash)) conjuntos.set(groupHash, []);
        conjuntos.get(groupHash).push(tr);
    });

    // Destacar duplicatas exatas (roxo claro)
    for (const [hash, rows] of hashes) {
        if (rows.length > 1) {
            rows.forEach(row => {
                row.style.backgroundColor = "#f0e6ff"; // Roxo claro
            });
        }
    }

    // Destacar duplicatas de conjunto (roxo médio)
    for (const [groupHash, rows] of conjuntos) {
        if (rows.length > 1) {
            // Apenas aplica se não for uma duplicata exata já colorida
            rows.forEach(row => {
                if (row.style.backgroundColor !== "#f0e6ff") {
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
    Swal.fire("✅ Lista iniciada", "10 linhas criadas com sucesso!", "success");
    verificarDuplicatas(); // Verifica duplicatas ao criar nova lista
});

document.getElementById("continuarListaBtn").addEventListener("click", () => {
    criar10Linhas();
    Swal.fire("➕ Adicionado", "10 novas linhas foram inseridas.", "success");
    verificarDuplicatas(); // Verifica duplicatas ao continuar lista
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
            // Se não houver linha suficiente para colar, cria uma nova
            criarLinha();
            targetRow = tabela.rows[index + i];
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
        tabela.innerHTML = ""; // Limpa a tabela existente

        json.forEach(rowData => {
            // Mapeamento de colunas do Excel para as propriedades internas
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
            criarLinha(mappedData); // Passa os dados mapeados para criar a linha
        });

        atualizarSequencias();
        verificarDuplicatas();
        Swal.fire("✅ Lista carregada", "A tabela foi preenchida com sucesso.", "success");
    } catch (error) {
        console.error("Erro ao carregar o arquivo:", error);
        Swal.fire("❌ Erro", "Não foi possível carregar a lista. Verifique o formato do arquivo.", "error");
    } finally {
        e.target.value = ''; // Limpa o input file para permitir carregar o mesmo arquivo novamente
    }
});


// Lógica para Exibir/Ocultar SEQ
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
            text: `Clique na célula ou linha que deseja pintar com ${corSelecionada}.`,
            icon: "info",
            timer: 2000,
            showConfirmButton: false
        });
    });
});

document.getElementById("paintFullRow").addEventListener("change", function() {
    pintarLinhaInteira = this.checked;
    if (pintarLinhaInteira) {
        Swal.fire("Pintar linha inteira ativado!", "Clique em um botão de cor e depois na linha para pintar.", "info");
    } else {
        Swal.fire("Pintar célula ativado!", "Clique em um botão de cor e depois na célula para pintar.", "info");
    }
});

// Lógica para selecionar/desselecionar todas as linhas
document.getElementById("toggleAllCheckboxes").addEventListener("change", function() {
    const checkboxes = document.querySelectorAll(".linha-selecao");
    checkboxes.forEach(cb => {
        cb.checked = this.checked;
    });
});

// Ao carregar a página, cria 10 linhas vazias para iniciar
document.addEventListener("DOMContentLoaded", () => {
    if (tabela.rows.length === 0) { // Cria linhas apenas se a tabela estiver vazia
        criar10Linhas();
        verificarDuplicatas();
    }
});