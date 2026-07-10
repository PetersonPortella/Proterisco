function mostrarSecao(id) {
    document.querySelectorAll(".secao").forEach(secao => {
        secao.classList.remove("ativa");
    });

    document.getElementById(id).classList.add("ativa");
}

let grafico1, grafico2, grafico3;

const estados = ["BA", "GO", "SP"];

document.getElementById("inputExcel").addEventListener("change", function (e) {
    const arquivo = e.target.files[0];
    const leitor = new FileReader();

    leitor.onload = function (evento) {
        const dados = new Uint8Array(evento.target.result);
        const workbook = XLSX.read(dados, { type: "array" });
        const aba = workbook.Sheets[workbook.SheetNames[0]];
        const linhas = XLSX.utils.sheet_to_json(aba);

        gerarGraficos(linhas);
    };

    leitor.readAsArrayBuffer(arquivo);
});

function gerarGraficos(linhas) {
    const baterias = { BA: 0, GO: 0, SP: 0 };
    const ocorrencias = { BA: 0, GO: 0, SP: 0 };

    const grupos = {
        "FURTO/ROUBO": 0,
        "VANDALISMO SEM FURTO/ROUBO": 0
    };

    let totalIds = 0;
    let totalBaterias = 0;
    let totalValor = 0;

    linhas.forEach(linha => {
        const estado = linha["Estado"];
        const id = linha["Id Ocorrência"];
        const grupo = linha["Grupo de Ocorrência"];
        const tipo = linha["Tipo de Prejuízo / Recuperação"];
        const quantidade = Number(linha["Quantidade"]) || 0;
        const valor = Number(linha["Valor Total"]) || 0;

        if (!estados.includes(estado)) return;

        if (id) {
            ocorrencias[estado]++;
            totalIds++;
        }

        if (tipo === "BATERIAS") {
            baterias[estado] += quantidade;
            totalBaterias += quantidade;
        }

        if (grupo === "FURTO/ROUBO") {
            grupos["FURTO/ROUBO"]++;
        }

        if (grupo === "VANDALISMO SEM FURTO/ROUBO") {
            grupos["VANDALISMO SEM FURTO/ROUBO"]++;
        }

        totalValor += valor;
    });

    criarGrafico1(baterias);
    criarGrafico2(ocorrencias);
    criarGrafico3(grupos);

    document.getElementById("kpiIds").innerText = totalIds;
    document.getElementById("kpiBaterias").innerText = totalBaterias;
    document.getElementById("kpiValor").innerText = totalValor.toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL"
    });
}

function criarGrafico1(dados) {
    if (grafico1) grafico1.destroy();

    grafico1 = new Chart(document.getElementById("grafico1"), {
        type: "bar",
        data: {
            labels: Object.keys(dados),
            datasets: [{
                data: Object.values(dados),
                backgroundColor: "#dc2626"
            }]
        },
        options: {
            indexAxis: "y",
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            }
        }
    });
}

function criarGrafico2(dados) {
    if (grafico2) grafico2.destroy();

    grafico2 = new Chart(document.getElementById("grafico2"), {
        type: "bar",
        data: {
            labels: Object.keys(dados),
            datasets: [{
                label: "Ocorrências",
                data: Object.values(dados),
                backgroundColor: "#dc2626"
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    });
}

function criarGrafico3(dados) {
    if (grafico3) grafico3.destroy();

    grafico3 = new Chart(document.getElementById("grafico3"), {
        type: "pie",
        data: {
            labels: Object.keys(dados),
            datasets: [{
                data: Object.values(dados),
                backgroundColor: ["#dc2626", "#475569"]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    });
}