const tabela = document.getElementById("tabelaLista");
const btnNovaLista = document.getElementById("btnNovaLista");
const btnSalvarLista = document.getElementById("btnSalvarLista");
const btnExportar = document.getElementById("btnExportar");
const btnCarregarLista = document.getElementById("btnCarregarLista");
const btnImportarLista = document.getElementById("btnImportarLista");
const btnResetarBanco = document.getElementById("btnResetarBanco");
const fileInput = document.getElementById("fileInput");

btnNovaLista.onclick = criarNovaLista;
btnSalvarLista.onclick = salvarLista;
btnExportar.onclick = exportarLista;
btnCarregarLista.onclick = carregarListaSalva;
btnImportarLista.onclick = () => fileInput.click();
btnResetarBanco.onclick = resetarBanco;

fileInput.addEventListener("change", importarArquivoExcel);

function criarNovaLista() {
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
  const rows = tabela.querySelectorAll("tbody tr");
  let sucesso = 0;
  for (const row of rows) {
    const cells = Array.from(row.querySelectorAll("input")).map(el => el.value.trim());
    if (cells.every(v => v === "")) continue;
    const body = {
      codigoMaterial: cells[0],
      nivel: cells[1],
      tipoEstrutura: cells[2],
      linha: cells[3],
      itemComponente: cells[4],
      qtdeMontagem: cells[5],
      unidadeMedida: cells[6],
      fatorSucata: cells[7]
    };
    const res = await fetch("/api/listas", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    if (res.ok) sucesso++;
  }
  Swal.fire("✅ Sucesso", `${sucesso} linhas salvas.`, "success");
}

async function carregarListaSalva() {
  try {
    const res = await fetch("/api/listas");
    const dados = await res.json();
    if (!dados.length) {
      Swal.fire("📂 Lista vazia", "Nenhum dado encontrado.", "info");
      return;
    }
    renderizarTabela(dados);
  } catch (e) {
    Swal.fire("Erro", "Falha ao carregar dados do banco.", "error");
  }
}

function importarArquivoExcel(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    renderizarTabela(json);
    Swal.fire("✅ Importado", "Lista carregada do Excel.", "success");
  };
  reader.readAsArrayBuffer(file);
}

async function resetarBanco() {
  const confirm = await Swal.fire({
    title: "Tem certeza?",
    text: "Todos os dados ser\u00e3o apagados.",
    icon: "warning",
    showCancelButton: true,
    confirmButtonText: "Sim, apagar",
    cancelButtonText: "Cancelar"
  });
  if (confirm.isConfirmed) {
    const res = await fetch("/api/listas/resetar", { method: "DELETE" });
    if (res.ok) {
      tabela.innerHTML = "";
      Swal.fire("🗑️ Banco Limpo", "Todos os dados foram apagados.", "success");
    } else {
      Swal.fire("Erro", "Falha ao resetar banco.", "error");
    }
  }
}

function renderizarTabela(dados) {
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
    headers.forEach(h => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.value = item[h] || "";
      td.appendChild(input);
      row.appendChild(td);
    });
    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  tabela.innerHTML = "";
  tabela.appendChild(table);
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
