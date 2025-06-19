let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];
let seqAtivo = true;
let seqCounter = 1;
let corSelecionada = "";
let pintarLinha = false;

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
    input.value = value.toUpperCase();
    input.addEventListener("input", (e) => {
        e.target.value = e.target.value.toUpperCase();
        verificarDuplicatas();
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
    seqTd.classList.add("seq");
    seqTd.textContent = seqCounter++;
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
}

function ajustarNivel(row, delta) {
    const select = row.querySelector(".nivel-select");
    let nivelAtual = parseInt(select.value);
    nivelAtual = Math.min(10, Math.max(1, nivelAtual + delta));
    select.value = nivelAtual;
    aplicarIndentacao(row);
}

function aplicarIndentacao(row) {
    for (let i = 1; i <= 10; i++) {
        row.classList.remove(`nivel-${i}`);
    }
    const nivel = row.querySelector(".nivel-select").value;
    row.classList.add(`nivel-${nivel}`);
}

function criar10Linhas() {
    for (let i = 0; i < 10; i++) criarLinha();
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
    cells[3].querySelector("select").value = data.NIVEL;
    aplicarIndentacao(row);
    cells[4].querySelector("select").value = data.SITE;
    cells[5].querySelector("select").value = data.ALTERNATIVA;
    cells[6].querySelector("input").value = data.CODIGO_MATERIAL;
    cells[7].querySelector("select").value = data.TIPO_ESTRUTURA;
    cells[8].querySelector("select").value = data.LINHA;
    cells[9].querySelector("input").value = data.ITEM_COMPONENTE;
    cells[10].querySelector("input").value = data.QTDE_MONTAGEM;
    cells[11].querySelector("select").value = data.UNIDADE_MEDIDA;
    cells[12].querySelector("select").value = data.FATOR_SUCATA;
}

function verificarDuplicatas() {
    const linhas = Array.from(tabela.rows);
    const hashes = new Map();
    const conjuntos = new Map();

    linhas.forEach(row => row.style.backgroundColor = "");

    linhas.forEach((tr) => {
        const data = getLinhaData(tr);
        if (data.CODIGO_MATERIAL === "" && data.ITEM_COMPONENTE === "") return;

        const hash = `${data.SITE}|${data.ALTERNATIVA}|${data.CODIGO_MATERIAL}|${data.TIPO_ESTRUTURA}|${data.ITEM_COMPONENTE}|${data.UNIDADE_MEDIDA}|${data.FATOR_SUCATA}`;
        const groupHash = `${data.CODIGO_MATERIAL}|${data.ITEM_COMPONENTE}`;

        if (!hashes.has(hash)) hashes.set(hash, []);
        hashes.get(hash).push({ row: tr, data });

        if (!conjuntos.has(groupHash)) conjuntos.set(groupHash, []);
        conjuntos.get(groupHash).push(tr);
    });

    for (const [, items] of hashes) {
        const ativos = items.filter(item => item.data.CODIGO_MATERIAL && item.data.ITEM_COMPONENTE);
        if (ativos.length > 1) {
            ativos.forEach(item => item.row.style.backgroundColor = "#f0e6ff");
        }
    }

    for (const [, rows] of conjuntos) {
        if (rows.length > 1) {
            rows.forEach(row => row.style.backgroundColor = "#d9c2ff");
        }
    }
}

document.getElementById("criarListaBtn").addEventListener("click", () => {
    tabela.innerHTML = "";
    seqCounter = 1;
    criar10Linhas();
    Swal.fire("✅ Lista iniciada", "10 linhas criadas com sucesso!", "success");
});

document.getElementById("continuarListaBtn").addEventListener("click", () => {
    criar10Linhas();
    Swal.fire("➕ Adicionado", "10 novas linhas foram inseridas.", "success");
});

document.getElementById("copiarSelecionadoBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length === 0) return Swal.fire("⚠️ Nada selecionado", "Marque uma linha.", "info");
    cacheCopiado = Array.from(selecionados).map(cb => getLinhaData(cb.closest("tr")));
    Swal.fire("📋 Copiado", `${cacheCopiado.length} linha(s) armazenada(s).`, "success");
});

document.getElementById("colarBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length !== 1) return Swal.fire("⚠️ Selecione uma única linha como referência.", "", "info");

    const trBase = selecionados[0].closest("tr");
    let index = Array.from(tabela.rows).indexOf(trBase);

    cacheCopiado.forEach((linha, i) => {
        let row = tabela.rows[index + i];
        if (!row) {
            criarLinha();
            row = tabela.rows[index + i];
        }
        preencherLinha(row, linha);
    });

    verificarDuplicatas();
});

document.getElementById("deletarSelecionadosBtn").addEventListener("click", () => {
    document.querySelectorAll(".linha-selecao:checked").forEach(cb => cb.closest("tr").remove());
});

document.getElementById("inserirAcimaBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length !== 1) return Swal.fire("⚠️ Selecione uma única linha", "Para inserir acima, selecione apenas uma.", "info");

    const row = selecionados[0].closest("tr");
    const nova = criarLinhaVazia();
    tabela.insertBefore(nova, row);
});

document.getElementById("inserirAbaixoBtn").addEventListener("click", () => {
    const selecionados = document.querySelectorAll(".linha-selecao:checked");
    if (selecionados.length !== 1) return Swal.fire("⚠️ Selecione uma única linha", "Para inserir abaixo, selecione apenas uma.", "info");

    const row = selecionados[0].closest("tr");
    const nova = criarLinhaVazia();
    tabela.insertBefore(nova, row.nextSibling);
});

document.getElementById("salvarListaBtn").addEventListener("click", () => {
    const dados = Array.from(tabela.rows).map(getLinhaData);
    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ListaIFS");
    XLSX.writeFile(wb, `lista_padrao_ifs_${new Date().toISOString().split("T")[0]}.xlsx`);
    Swal.fire("💾 Lista Exportada", "Arquivo Excel gerado com sucesso.", "success");
});

document.getElementById("inputFile").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    tabela.innerHTML = "";
    seqCounter = 1;
    json.forEach(row => criarLinha({
        SITE: row.SITE,
        ALTERNATIVA: row.ALTERNATIVA,
        CODIGO_MATERIAL: row["CÓDIGO_MATERIAL"] || row["CODIGO MATERIAL"],
        NIVEL: row.NIVEL,
        TIPO_ESTRUTURA: row["TIPO ESTRUTURA"],
        LINHA: row.LINHA,
        ITEM_COMPONENTE: row.ITEM_COMPONENTE,
        QTDE_MONTAGEM: row.QTDE_MONTAGEM,
        UNIDADE_MEDIDA: row["UNIDADE DE MEDIDA"],
        FATOR_SUCATA: row.FATOR_SUCATA,
    }));
    Swal.fire("✅ Lista carregada", "A tabela foi preenchida com sucesso.", "success");
});

function criarLinhaVazia() {
    const dummy = document.createElement("tbody");
    criarLinha();
    dummy.appendChild(tabela.lastChild);
    return dummy.removeChild(dummy.firstChild);
}

// Filtro por nível
document.getElementById("filtroNivel").addEventListener("change", (e) => {
    const nivelFiltro = e.target.value;
    Array.from(tabela.rows).forEach(row => {
        const nivel = row.querySelector(".nivel-select").value;
        row.style.display = (!nivelFiltro || nivel === nivelFiltro) ? "" : "none";
    });
});

// Toggle coluna SEQ
document.getElementById("toggleSeqBtn").addEventListener("click", () => {
    const thSeq = document.querySelector(".seq-col");
    const tdSeqs = document.querySelectorAll("td.seq");
    seqAtivo = !seqAtivo;
    thSeq.classList.toggle("hidden", !seqAtivo);
    tdSeqs.forEach(td => td.classList.toggle("hidden", !seqAtivo));
});

// Pintura com botão de cor
document.querySelectorAll(".paint-btn").forEach(btn => {
    btn.addEventListener("click", () => {
        corSelecionada = btn.dataset.color;
        pintarLinha = document.getElementById("paintFullRow").checked;
        Swal.fire("🎨 Modo pintura ativo", "Clique sobre uma célula ou linha.", "info");
    });
});

document.addEventListener("click", (e) => {
    if (!corSelecionada) return;
    if (e.target.tagName === "TD") {
        if (pintarLinha) {
            e.target.closest("tr").style.backgroundColor = corSelecionada;
        } else {
            e.target.style.backgroundColor = corSelecionada;
        }
        corSelecionada = "";
    }
});
