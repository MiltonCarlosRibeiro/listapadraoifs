let listaAtual = [];

const tabela = document.getElementById("tabelaLista");
const btnNovaLista = document.getElementById("btnNovaLista");
const btnSalvarLista = document.getElementById("btnSalvarLista");
const btnExportar = document.getElementById("btnExportar");
const btnCarregarLista = document.getElementById("btnCarregarLista");

btnNovaLista.onclick = criarNovaLista;
btnSalvarLista.onclick = salvarLista;
btnExportar.onclick = exportarLista;
if (btnCarregarLista) btnCarregarLista.onclick = carregarListaSalva;

function criarNovaLista() {
  listaAtual = [];
  const headers = [
    "codigoMaterial", "nivel", "tipoEstrutura",
    "linha", "itemComponente", "qtdeMontagem",
    "unidadeMedida", "fatorSucata"
  ];

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let i = 0; i < 10; i++) {
    const row = document.createElement("tr");
    headers.forEach(() => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      td.appendChild(input);
      row.appendChild(td);
    });
    tbody.appendChild(row);
  }
  table.appendChild(tbody);

  tabela.innerHTML = "";
  tabela.appendChild(table);
}

async function salvarLista() {
  const inputs = tabela.querySelectorAll("tbody tr");
  let sucesso = 0;
  for (const row of inputs) {
    const dados = Array.from(row.querySelectorAll("input")).map(el => el.value.trim());
    if (dados.every(v => v === "")) continue; // ignora linha vazia

    const body = {
      codigoMaterial: dados[0],
      nivel: dados[1],
      tipoEstrutura: dados[2],
      linha: dados[3],
      itemComponente: dados[4],
      qtdeMontagem: dados[5],
      unidadeMedida: dados[6],
      fatorSucata: dados[7]
    };

    const res = await fetch("/api/listas", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    if (res.ok) sucesso++;
  }

  Swal.fire("✅ Sucesso", `${sucesso} linha(s) salvas com sucesso!`, "success");
}

async function carregarListaSalva() {
  const res = await fetch("/api/listas");
  const dados = await res.json();

  if (!dados.length) return Swal.fire("ℹ️ Aviso", "Nenhuma lista salva encontrada.", "info");

  const headers = [
    "codigoMaterial", "nivel", "tipoEstrutura",
    "linha", "itemComponente", "qtdeMontagem",
    "unidadeMedida", "fatorSucata"
  ];

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  dados.forEach(item => {
    const row = document.createElement("tr");
    headers.forEach(key => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.value = item[key] || "";
      td.appendChild(input);
      row.appendChild(td);
    });
    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  tabela.innerHTML = "";
  tabela.appendChild(table);

  Swal.fire("✅ Lista carregada!", `Foram carregadas ${dados.length} linha(s).`, "success");
}

function exportarLista() {
  const linhas = tabela.querySelectorAll("tbody tr");
  const dados = [];
  linhas.forEach(row => {
    const cells = Array.from(row.querySelectorAll("input")).map(el => el.value.trim());
    if (cells.every(v => v === "")) return;
    dados.push({
      codigoMaterial: cells[0],
      nivel: cells[1],
      tipoEstrutura: cells[2],
      linha: cells[3],
      itemComponente: cells[4],
      qtdeMontagem: cells[5],
      unidadeMedida: cells[6],
      fatorSucata: cells[7]
    });
  });

  const ws = XLSX.utils.json_to_sheet(dados);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "ListaIFS");
  XLSX.writeFile(wb, "lista-padrao-ifs.xlsx");
}