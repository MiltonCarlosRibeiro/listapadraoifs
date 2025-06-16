// script.js FINALIZADO E FUNCIONAL

let tabela = document.getElementById("listaTabela").getElementsByTagName("tbody")[0];
let cacheCopiado = [];

const unidades = ["un", "cj", "kg", "mm", "m"];
const tiposEstrutura = ["Manufatura", "Comprado", ""];
const fatorSucata = ["0", "15", ""];
const linhaValores = Array.from({ length: 80 }, (_, i) => String((i + 1) * 10));
const alternativas = ["*", ""];
const siteValores = ["1", ""];

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

function selectCell(options = [], selected = "") {
    const td = document.createElement("td");
    const select = document.createElement("select");
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

    row.appendChild(selectCell(siteValores, v.SITE || "1"));
    row.appendChild(selectCell(alternativas, v.ALTERNATIVA || "*"));
    row.appendChild(inputCell("text", false, v.CODIGO_MATERIAL || ""));
    row.appendChild(inputCell("text", false, v.NIVEL || "1"));
    row.appendChild(selectCell(tiposEstrutura, v.TIPO_ESTRUTURA || "Manufatura"));
    row.appendChild(selectCell(linhaValores, v.LINHA || "10"));
    row.appendChild(inputCell("text", false, v.ITEM_COMPONENTE || ""));
    row.appendChild(inputCell("number", false, v.QTDE_MONTAGEM || "0"));
    row.appendChild(selectCell(unidades, v.UNIDADE_MEDIDA || "un"));
    row.appendChild(selectCell(fatorSucata, v.FATOR_SUCATA || "0"));

    tabela.appendChild(row);
}

function criar10Linhas() {
    for (let i = 0; i < 10; i++) criarLinha();
}

function getLinhaData(tr) {
    const cells = tr.querySelectorAll("td");
    return {
        SITE: cells[1]?.querySelector("select")?.value || "",
        ALTERNATIVA: cells[2]?.querySelector("select")?.value || "",
        CODIGO_MATERIAL: cells[3]?.querySelector("input")?.value.trim().toUpperCase() || "",
        NIVEL: cells[4]?.querySelector("input")?.value.trim() || "",
        TIPO_ESTRUTURA: cells[5]?.querySelector("select")?.value || "",
        LINHA: cells[6]?.querySelector("select")?.value || "",
        ITEM_COMPONENTE: cells[7]?.querySelector("input")?.value.trim().toUpperCase() || "",
        QTDE_MONTAGEM: cells[8]?.querySelector("input")?.value.trim() || "",
        UNIDADE_MEDIDA: cells[9]?.querySelector("select")?.value || "",
        FATOR_SUCATA: cells[10]?.querySelector("select")?.value || ""
    };
}

function preencherLinha(row, data) {
    const cells = row.querySelectorAll("td");
    cells[1].querySelector("select").value = data.SITE;
    cells[2].querySelector("select").value = data.ALTERNATIVA;
    cells[3].querySelector("input").value = data.CODIGO_MATERIAL;
    cells[4].querySelector("input").value = data.NIVEL;
    cells[5].querySelector("select").value = data.TIPO_ESTRUTURA;
    cells[6].querySelector("select").value = data.LINHA;
    cells[7].querySelector("input").value = data.ITEM_COMPONENTE;
    cells[8].querySelector("input").value = data.QTDE_MONTAGEM;
    cells[9].querySelector("select").value = data.UNIDADE_MEDIDA;
    cells[10].querySelector("select").value = data.FATOR_SUCATA;
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

// Ações de botão
document.getElementById("criarListaBtn").addEventListener("click", () => {
    tabela.innerHTML = "";
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
    const dados = Array.from(tabela.rows).map(tr => {
        const cells = tr.querySelectorAll("td");
        return {
            SITE: cells[1].querySelector("select").value,
            ALTERNATIVA: cells[2].querySelector("select").value,
            CÓDIGO_MATERIAL: cells[3].querySelector("input").value.trim().toUpperCase(),
            NÍVEL: cells[4].querySelector("input").value.trim(),
            TIPO_ESTRUTURA: cells[5].querySelector("select").value,
            LINHA: cells[6].querySelector("select").value,
            ITEM_COMPONENTE: cells[7].querySelector("input").value.trim().toUpperCase(),
            QTDE_MONTAGEM: cells[8].querySelector("input").value.trim(),
            UNIDADE_DE_MEDIDA: cells[9].querySelector("select").value,
            FATOR_SUCATA: cells[10].querySelector("select").value
        };
    });

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
