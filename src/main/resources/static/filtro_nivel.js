function aplicarFiltroNiveis() {
    const selecionados = Array.from(document.querySelectorAll("#checkboxNiveis input:checked")).map(cb => cb.value);
    const linhas = document.querySelectorAll("#listaTabela tbody tr");

    linhas.forEach(tr => {
        // Agora o NÍVEL é um input de texto, então selecionamos o input dentro da célula
        const nivelInput = tr.querySelector(".nivel-col input");
        const nivel = nivelInput ? nivelInput.value : "";

        if (selecionados.length === 0 || selecionados.includes(nivel)) {
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
        observer.observe(target, { childList: true, subtree: true });
    }
    configurarFiltroNiveis();
    aplicarFiltroNiveis(); // Aplica o filtro inicial ao carregar
}

document.addEventListener("DOMContentLoaded", () => {
    ativarFiltroAuto();
    // Adiciona listener para input/change em todas as células de Nível existentes
    // Isso é importante para que o filtro seja re-aplicado quando o NÍVEL for alterado
    // em uma célula (seja por digitação ou colagem)
    document.querySelectorAll(".nivel-col input").forEach(input => {
        input.addEventListener('input', aplicarFiltroNiveis);
    });
});