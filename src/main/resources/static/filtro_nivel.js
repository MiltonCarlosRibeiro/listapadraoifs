function aplicarFiltroNiveis() {
    const selecionados = Array.from(document.querySelectorAll("#checkboxNiveis input:checked")).map(cb => cb.value);
    const linhas = document.querySelectorAll("#listaTabela tbody tr");

    linhas.forEach(tr => {
        // Pega o valor do input dentro da célula de nível
        const nivelInput = tr.querySelector("td.nivel-col input[type='text']");
        const nivel = nivelInput ? nivelInput.value : ''; // Garante que pegamos o valor do input

        if (selecionados.length === 0 || selecionados.includes(nivel)) {
            tr.style.display = ""; // Exibe a linha
        } else {
            tr.style.display = "none"; // Oculta a linha
        }
    });
}

function configurarFiltroNiveis() {
    const checkboxesContainer = document.getElementById("checkboxNiveis");
    checkboxesContainer.innerHTML = ''; // Limpa antes de gerar

    // Gera os checkboxes dinamicamente
    for (let i = 1; i <= 10; i++) {
        const label = document.createElement("label");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = String(i);
        checkbox.id = `filterNivel${i}`;
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(`Nível ${i}`));
        checkboxesContainer.appendChild(label);
    }

    const checkboxes = document.querySelectorAll("#checkboxNiveis input[type=checkbox]");
    checkboxes.forEach(cb => cb.removeEventListener("change", aplicarFiltroNiveis)); // Remove para evitar duplicidade
    checkboxes.forEach(cb => cb.addEventListener("change", aplicarFiltroNiveis));

    document.getElementById("limparFiltroNivelBtn").addEventListener("click", () => {
        checkboxes.forEach(cb => cb.checked = false); // Desmarca todos
        aplicarFiltroNiveis(); // Aplica o filtro sem seleções, mostrando tudo
    });

    // Adiciona listener para mudanças nos inputs de nível (do script.js)
    document.getElementById("listaTabela").addEventListener('change', function(event) {
        if (event.target.closest('td.nivel-col')) { // Se a mudança ocorreu em uma célula de nível
            aplicarFiltroNiveis(); // Re-aplica o filtro
        }
    });
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