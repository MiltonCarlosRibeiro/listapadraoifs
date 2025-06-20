function aplicarFiltroNiveis() {
    const selecionados = Array.from(document.querySelectorAll("#checkboxNiveis input:checked")).map(cb => cb.value);
    const linhas = document.querySelectorAll("#listaTabela tbody tr");

    linhas.forEach(tr => {
        // A coluna NÍVEL é agora a TD de índice 2 na linha
        const nivelInput = tr.querySelectorAll("td")[2]?.querySelector("input");
        const nivel = nivelInput ? nivelInput.value : null;

        if (selecionados.length === 0 || (nivel && selecionados.includes(nivel))) {
            tr.style.display = ""; // Exibe a linha
        } else {
            tr.style.display = "none"; // Oculta a linha
        }
    });
}

function configurarFiltroNiveis() {
    const checkboxes = document.querySelectorAll("#checkboxNiveis input[type=checkbox]");
    checkboxes.forEach(cb => cb.removeEventListener("change", aplicarFiltroNiveis)); // Remove para evitar duplicidade
    checkboxes.forEach(cb => cb.addEventListener("change", aplicarFiltroNiveis));

    document.getElementById("limparFiltroNivelBtn").addEventListener("click", () => {
        checkboxes.forEach(cb => cb.checked = false); // Desmarca todos
        aplicarFiltroNiveis(); // Aplica o filtro sem seleções, mostrando tudo
    });
}

function ativarFiltroAuto() {
    const observer = new MutationObserver(aplicarFiltroNiveis);
    const target = document.querySelector("#listaTabela tbody");
    if (target) {
        // Observa mudanças de filhos (linhas adicionadas/removidas) e mudanças em subárvores (inputs nas células)
        observer.observe(target, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] });
    }
    configurarFiltroNiveis();
}

document.addEventListener("DOMContentLoaded", () => {
    ativarFiltroAuto();
    // Re-aplicar filtro quando o script.js adiciona linhas iniciais
    // Ou quando NÍVEL é modificado (evento 'change' buble)
    document.getElementById("listaTabela").addEventListener('change', (event) => {
        if (event.target.closest('.nivel-col')) {
            aplicarFiltroNiveis();
        }
    });
});