/**
 * @file filtro_nivel.js
 * @description Gerencia a funcionalidade de filtrar as linhas da tabela pelo seu NÍVEL.
 */

function aplicarFiltroNiveis() {
    const selecionados = Array.from(document.querySelectorAll("#checkboxNiveis input:checked")).map(cb => cb.value);
    const linhas = document.querySelectorAll("#listaTabela tbody tr");

    linhas.forEach(tr => {
        const nivelInput = tr.querySelector("td.nivel-col input[type='text']");
        const nivel = nivelInput ? nivelInput.value : '';

        if (selecionados.length === 0 || selecionados.includes(nivel)) {
            tr.style.display = "";
        } else {
            tr.style.display = "none";
        }
    });
}

function configurarFiltroNiveis() {
    const checkboxesContainer = document.getElementById("checkboxNiveis");
    if (!checkboxesContainer) return;
    checkboxesContainer.innerHTML = '';

    for (let i = 1; i <= 10; i++) {
        const label = document.createElement("label");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = String(i);
        checkbox.id = `filterNivel${i}`;
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(` Nível ${i}`));
        checkboxesContainer.appendChild(label);
    }

    const checkboxes = document.querySelectorAll("#checkboxNiveis input[type=checkbox]");
    checkboxes.forEach(cb => {
        cb.removeEventListener("change", aplicarFiltroNiveis);
        cb.addEventListener("change", aplicarFiltroNiveis);
    });

    const limparFiltroBtn = document.getElementById("limparFiltroNivelBtn");
    if(limparFiltroBtn) {
        limparFiltroBtn.addEventListener("click", () => {
            checkboxes.forEach(cb => cb.checked = false);
            aplicarFiltroNiveis();
        });
    }

    const tabelaElement = document.getElementById("listaTabela");
    if(tabelaElement) {
        tabelaElement.addEventListener('change', function(event) {
            if (event.target.closest('td.nivel-col')) {
                aplicarFiltroNiveis();
            }
        });
    }
}

function ativarFiltroAuto() {
    const observer = new MutationObserver(aplicarFiltroNiveis);
    const target = document.querySelector("#listaTabela tbody");
    if (target) {
        observer.observe(target, { childList: true, subtree: true });
    }
    configurarFiltroNiveis();
}

document.addEventListener("DOMContentLoaded", () => {
    ativarFiltroAuto();
});
// A CHAVE "}" EXTRA QUE ESTAVA AQUI FOI REMOVIDA