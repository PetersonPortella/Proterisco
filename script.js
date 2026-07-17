/* =========================================================
   NAVEGAÇÃO ENTRE AS ABAS
========================================================= */

function mostrarSecao(id) {
    document.querySelectorAll(".secao").forEach(secao => {
        secao.classList.remove("ativa");
    });

    const secaoSelecionada = document.getElementById(id);

    if (secaoSelecionada) {
        secaoSelecionada.classList.add("ativa");
    }
}

let graficoBateriasMensalGoias = null;

let graficoTop10SitesBahia = null;
/* =========================================================
   CONFIGURAÇÕES GERAIS
========================================================= */

const caminhoPlanilhaMestra = "./dados/SIGO_2026.xlsx";

const estadosMonitorados = [
    "BAHIA",
    "GOIAS",
    "SAO PAULO"
];

const aliasesCabecalhos = {
    "tipos de prejuizo / recuperacao":
        "item de prejuizo / recuperacao",

    "tipo de prejuizo / recuperacao":
        "item de prejuizo / recuperacao",

    "valor":
        "valor total"
};

let workbookAtualizado = null;
let linhasMestra = [];
let linhasNovas = [];
let baseSIGO2026 = [];
let cabecalhosMestra = [];

let periodoSelecionado = "semana";
let dataReferencia = new Date();

let graficoBateriasCidades = null;
let graficoCabosCidades = null;
let graficoOcorrenciasEstados = null;


/* =========================================================
   ELEMENTOS DO HTML
========================================================= */

const inputExcel =
    document.getElementById("inputExcel");

const salvarPlanilha =
    document.getElementById("salvarPlanilha");

const statusMestra =
    document.getElementById("statusMestra");

const resultadoImportacao =
    document.getElementById("resultadoImportacao");

const periodoSelecionadoTexto =
    document.getElementById("periodoSelecionadoTexto");


/* =========================================================
   PLUGIN PARA MOSTRAR NÚMEROS NOS GRÁFICOS
========================================================= */

const pluginValores = {
    id: "pluginValores",

    afterDatasetsDraw(chart) {
        const contexto = chart.ctx;

        contexto.save();
        contexto.fillStyle = "#f8fafc";
        contexto.font = "bold 12px Arial";
        contexto.textAlign = "center";
        contexto.textBaseline = "bottom";

        chart.data.datasets.forEach(
            (dataset, indiceDataset) => {
                const meta =
                    chart.getDatasetMeta(indiceDataset);

                if (meta.hidden) return;

                meta.data.forEach(
                    (elemento, indice) => {
                        const valor =
                            Number(dataset.data[indice]) || 0;

                        if (valor === 0) return;

                        const posicao =
                            elemento.tooltipPosition();

                        if (
                            chart.options.indexAxis === "y"
                        ) {
                            contexto.textAlign = "left";
                            contexto.textBaseline = "middle";

                            contexto.fillText(
                                formatarNumero(valor),
                                posicao.x + 8,
                                posicao.y
                            );
                        } else {
                            contexto.textAlign = "center";
                            contexto.textBaseline = "bottom";

                            contexto.fillText(
                                formatarNumero(valor),
                                posicao.x,
                                posicao.y - 8
                            );
                        }
                    }
                );
            }
        );

        contexto.restore();
    }
};

Chart.register(pluginValores);


/* =========================================================
   INICIALIZAÇÃO
========================================================= */

configurarBotoesPeriodo();
configurarAnimacoesScroll();
carregarPlanilhaMestra();


/* =========================================================
   BOTÕES ANO, MÊS E SEMANA
========================================================= */

function configurarBotoesPeriodo() {
    const botoes =
        document.querySelectorAll(".botao-periodo");

    botoes.forEach(botao => {
        botao.addEventListener("click", function () {
            botoes.forEach(item => {
                item.classList.remove("ativo");
            });

            this.classList.add("ativo");

            periodoSelecionado =
                this.dataset.periodo || "semana";

            atualizarDashboard();
        });
    });
}


/* =========================================================
   CARREGAR PLANILHA MESTRA
========================================================= */

async function carregarPlanilhaMestra() {
    try {
        statusMestra.textContent =
            "Carregando planilha mestra...";

        const resposta =
            await fetch(caminhoPlanilhaMestra);

        if (!resposta.ok) {
            throw new Error(
                "Planilha mestra não encontrada."
            );
        }

        const arquivo =
            await resposta.arrayBuffer();

        workbookAtualizado = XLSX.read(
            arquivo,
            {
                type: "array",
                cellDates: true
            }
        );

        const nomeAba =
            workbookAtualizado.SheetNames[0];

        const abaMestra =
            workbookAtualizado.Sheets[nomeAba];

        const dadosMestra =
            XLSX.utils.sheet_to_json(
                abaMestra,
                {
                    header: 1,
                    defval: "",
                    raw: true
                }
            );

        const indiceCabecalho =
            localizarLinhaCabecalho(dadosMestra);

        if (indiceCabecalho === -1) {
            throw new Error(
                "Não foi possível encontrar os cabeçalhos da planilha mestra."
            );
        }

        cabecalhosMestra =
            dadosMestra[indiceCabecalho]
                .map(valor =>
                    String(valor ?? "").trim()
                )
                .filter(cabecalho =>
                    cabecalho !== ""
                );

        linhasMestra = dadosMestra
            .slice(indiceCabecalho + 1)
            .filter(linha =>
                linha.some(valor =>
                    String(valor ?? "").trim() !== ""
                )
            )
            .map(linha =>
                transformarLinhaEmObjeto(
                    linha,
                    cabecalhosMestra
                )
            );

        baseSIGO2026 = [...linhasMestra];

        dataReferencia =
            encontrarDataReferencia(baseSIGO2026);

        statusMestra.textContent =
            `Planilha mestra carregada: ` +
            `${formatarNumero(linhasMestra.length)} linhas.`;

        resultadoImportacao.textContent = "";
        salvarPlanilha.disabled = true;

        atualizarDashboard();

    } catch (erro) {
        console.error(erro);

        statusMestra.textContent =
            "Erro ao carregar dados/SIGO_2026.xlsx.";
    }
}


/* =========================================================
   IMPORTAR PLANILHA SEMANAL
========================================================= */

inputExcel.addEventListener(
    "change",
    async function (evento) {
        const arquivo =
            evento.target.files[0];

        resultadoImportacao.textContent = "";
        salvarPlanilha.disabled = true;

        if (!arquivo) return;

        if (!workbookAtualizado) {
            resultadoImportacao.textContent =
                "A planilha mestra ainda não foi carregada.";

            return;
        }

        try {
            linhasNovas =
                await lerPlanilhaSemanal(arquivo);

            if (linhasNovas.length === 0) {
                throw new Error(
                    "A planilha semanal não possui dados."
                );
            }

            validarCabecalhos(
                cabecalhosMestra,
                linhasNovas
            );

            linhasNovas =
                padronizarLinhasNovas(
                    linhasNovas,
                    cabecalhosMestra
                );

            baseSIGO2026 = [
                ...baseSIGO2026,
                ...linhasNovas
            ];

            atualizarWorkbook();

            dataReferencia =
                encontrarDataReferencia(baseSIGO2026);

            resultadoImportacao.textContent =
                `${formatarNumero(linhasNovas.length)} ` +
                `linhas incorporadas. Novo total: ` +
                `${formatarNumero(baseSIGO2026.length)} linhas.`;

            salvarPlanilha.disabled = false;

            atualizarDashboard();

        } catch (erro) {
            console.error(erro);

            resultadoImportacao.textContent =
                erro.message;

            salvarPlanilha.disabled = true;
        }
    }
);


/* =========================================================
   LEITURA DA PLANILHA SEMANAL
========================================================= */

async function lerPlanilhaSemanal(arquivo) {
    const conteudo =
        await arquivo.arrayBuffer();

    if (
        arquivo.name.toLowerCase().endsWith(".xls") &&
        arquivoEhHtml(conteudo)
    ) {
        return lerArquivoHtmlXls(conteudo);
    }

    const workbookNovo = XLSX.read(
        conteudo,
        {
            type: "array",
            cellDates: true
        }
    );

    const nomeAbaNova =
        workbookNovo.SheetNames[0];

    const abaNova =
        workbookNovo.Sheets[nomeAbaNova];

    return lerAbaExcel(abaNova);
}


function arquivoEhHtml(conteudo) {
    const bytes =
        new Uint8Array(
            conteudo.slice(0, 300)
        );

    const inicio =
        new TextDecoder("windows-1252")
            .decode(bytes);

    const texto =
        inicio.trim().toLowerCase();

    return (
        texto.startsWith("<!doctype html") ||
        texto.startsWith("<html") ||
        texto.includes("<table")
    );
}


function lerArquivoHtmlXls(conteudo) {
    const texto =
        new TextDecoder("windows-1252")
            .decode(conteudo);

    const documento =
        new DOMParser().parseFromString(
            texto,
            "text/html"
        );

    const tabela =
        documento.querySelector("table");

    if (!tabela) {
        throw new Error(
            "Nenhuma tabela foi encontrada no arquivo XLS."
        );
    }

    const linhasTabela =
        Array.from(
            tabela.querySelectorAll("tr")
        );

    if (linhasTabela.length < 2) {
        throw new Error(
            "A planilha semanal não possui dados."
        );
    }

    const indiceCabecalho =
        localizarCabecalhoTabelaHtml(
            linhasTabela
        );

    if (indiceCabecalho === -1) {
        throw new Error(
            "Não foi possível localizar o cabeçalho do arquivo XLS."
        );
    }

    const cabecalhos =
        Array.from(
            linhasTabela[indiceCabecalho]
                .querySelectorAll("th, td")
        ).map(celula =>
            limparTexto(celula.textContent)
        );

    return linhasTabela
        .slice(indiceCabecalho + 1)
        .map(linha => {
            const celulas =
                Array.from(
                    linha.querySelectorAll(
                        "th, td"
                    )
                );

            const objeto = {};

            cabecalhos.forEach(
                (cabecalho, indice) => {
                    if (!cabecalho) return;

                    objeto[cabecalho] =
                        limparTexto(
                            celulas[indice]
                                ?.textContent ?? ""
                        );
                }
            );

            return objeto;
        })
        .filter(linha =>
            Object.values(linha).some(
                valor =>
                    String(valor ?? "").trim() !== ""
            )
        );
}


function localizarCabecalhoTabelaHtml(
    linhas
) {
    return linhas.findIndex(linha => {
        const valores =
            Array.from(
                linha.querySelectorAll(
                    "th, td"
                )
            ).map(celula =>
                normalizarCabecalho(
                    celula.textContent
                )
            );

        const esperados = [
            "estado",
            "grupo de ocorrencia",
            "quantidade"
        ];

        const encontrados =
            esperados.filter(cabecalho =>
                valores.includes(cabecalho)
            );

        return encontrados.length >= 2;
    });
}


function lerAbaExcel(aba) {
    const dados =
        XLSX.utils.sheet_to_json(
            aba,
            {
                header: 1,
                defval: "",
                raw: true
            }
        );

    const indiceCabecalho =
        localizarLinhaCabecalho(dados);

    if (indiceCabecalho === -1) {
        throw new Error(
            "Não foi possível localizar o cabeçalho da planilha."
        );
    }

    const cabecalhos =
        dados[indiceCabecalho].map(valor =>
            String(valor ?? "").trim()
        );

    return dados
        .slice(indiceCabecalho + 1)
        .filter(linha =>
            linha.some(valor =>
                String(valor ?? "").trim() !== ""
            )
        )
        .map(linha =>
            transformarLinhaEmObjeto(
                linha,
                cabecalhos
            )
        );
}


function localizarLinhaCabecalho(dados) {
    const cabecalhosEsperados = [
        "id ocorrencia",
        "estado",
        "grupo de ocorrencia",
        "quantidade"
    ];

    return dados.findIndex(linha => {
        const valoresNormalizados =
            linha.map(normalizarCabecalho);

        const encontrados =
            cabecalhosEsperados.filter(
                cabecalho =>
                    valoresNormalizados.includes(
                        cabecalho
                    )
            );

        return encontrados.length >= 3;
    });
}


function transformarLinhaEmObjeto(
    linha,
    cabecalhos
) {
    const objeto = {};

    cabecalhos.forEach(
        (cabecalho, indice) => {
            if (!cabecalho) return;

            objeto[cabecalho] =
                linha[indice] ?? "";
        }
    );

    return objeto;
}


/* =========================================================
   VALIDAÇÃO DOS CABEÇALHOS
========================================================= */

function validarCabecalhos(
    cabecalhosDaMestra,
    novasLinhas
) {
    if (novasLinhas.length === 0) {
        throw new Error(
            "A planilha selecionada está vazia."
        );
    }

    const cabecalhosNova =
        Object.keys(novasLinhas[0]);

    const cabecalhosNovaNormalizados =
        cabecalhosNova.map(
            normalizarCabecalho
        );

    const faltando =
        cabecalhosDaMestra.filter(
            cabecalhoMestra => {
                const normalizado =
                    normalizarCabecalho(
                        cabecalhoMestra
                    );

                return !cabecalhosNovaNormalizados
                    .includes(normalizado);
            }
        );

    if (faltando.length > 0) {
        throw new Error(
            `Planilha inválida. Colunas ausentes: ` +
            faltando.join(", ")
        );
    }
}


function padronizarLinhasNovas(
    novasLinhas,
    cabecalhosDaMestra
) {
    return novasLinhas.map(linha => {
        const linhaPadronizada = {};
        const mapaNovaLinha = {};

        Object.keys(linha).forEach(
            cabecalho => {
                mapaNovaLinha[
                    normalizarCabecalho(
                        cabecalho
                    )
                ] = linha[cabecalho];
            }
        );

        cabecalhosDaMestra.forEach(
            cabecalhoMestra => {
                const chaveNormalizada =
                    normalizarCabecalho(
                        cabecalhoMestra
                    );

                linhaPadronizada[
                    cabecalhoMestra
                ] =
                    mapaNovaLinha[
                        chaveNormalizada
                    ] ?? "";
            }
        );

        return linhaPadronizada;
    });
}


/* =========================================================
   ATUALIZAR E SALVAR EXCEL
========================================================= */

function atualizarWorkbook() {
    const nomeAba =
        workbookAtualizado.SheetNames[0];

    const dadosParaSalvar = [
        cabecalhosMestra,

        ...baseSIGO2026.map(linha =>
            cabecalhosMestra.map(
                cabecalho =>
                    linha[cabecalho] ?? ""
            )
        )
    ];

    const novaAba =
        XLSX.utils.aoa_to_sheet(
            dadosParaSalvar
        );

    workbookAtualizado.Sheets[nomeAba] =
        novaAba;
}


salvarPlanilha.addEventListener(
    "click",
    function () {
        if (!workbookAtualizado) return;

        XLSX.writeFile(
            workbookAtualizado,
            "SIGO_2026_ATUALIZADA.xlsx"
        );
    }
);


/* =========================================================
   ATUALIZAÇÃO COMPLETA DO DASHBOARD
========================================================= */

function atualizarDashboard() {
    if (!baseSIGO2026.length) return;

    const contexto =
        criarContextoPeriodo(
            baseSIGO2026
        );

    atualizarTextoPeriodo(contexto);

    atualizarBlocoBaterias(contexto);
    atualizarBlocoCabos(contexto);
    atualizarOcorrenciasEstados(contexto);
    atualizarResumoFinanceiro(contexto);
}


/* =========================================================
   PERÍODOS
========================================================= */

function criarContextoPeriodo(base) {
    const referencia =
        new Date(dataReferencia);

    const periodoAtual =
        obterIntervaloAtual(
            periodoSelecionado,
            referencia
        );

    const periodoAnterior =
        obterIntervaloAnterior(
            periodoSelecionado,
            periodoAtual
        );

    return {
        referencia,
        periodoAtual,
        periodoAnterior,

        linhasAtual:
            filtrarPorIntervalo(
                base,
                periodoAtual.inicio,
                periodoAtual.fim
            ),

        linhasAnterior:
            filtrarPorIntervalo(
                base,
                periodoAnterior.inicio,
                periodoAnterior.fim
            )
    };
}


function obterIntervaloAtual(
    tipo,
    referencia
) {
    if (tipo === "ano") {
        return {
            inicio: new Date(
                referencia.getFullYear(),
                0,
                1
            ),

            fim: new Date(
                referencia.getFullYear(),
                11,
                31,
                23,
                59,
                59,
                999
            )
        };
    }

    if (tipo === "mes") {
        return {
            inicio: new Date(
                referencia.getFullYear(),
                referencia.getMonth(),
                1
            ),

            fim: new Date(
                referencia.getFullYear(),
                referencia.getMonth() + 1,
                0,
                23,
                59,
                59,
                999
            )
        };
    }

    return obterSemanaDaData(referencia);
}


function obterIntervaloAnterior(
    tipo,
    intervaloAtual
) {
    if (tipo === "ano") {
        const ano =
            intervaloAtual.inicio
                .getFullYear() - 1;

        return {
            inicio: new Date(ano, 0, 1),

            fim: new Date(
                ano,
                11,
                31,
                23,
                59,
                59,
                999
            )
        };
    }

    if (tipo === "mes") {
        const inicio =
            new Date(
                intervaloAtual.inicio
                    .getFullYear(),
                intervaloAtual.inicio
                    .getMonth() - 1,
                1
            );

        return {
            inicio,

            fim: new Date(
                inicio.getFullYear(),
                inicio.getMonth() + 1,
                0,
                23,
                59,
                59,
                999
            )
        };
    }

    const fim =
        new Date(intervaloAtual.inicio);

    fim.setDate(
        fim.getDate() - 1
    );

    fim.setHours(
        23,
        59,
        59,
        999
    );

    const inicio =
        new Date(fim);

    inicio.setDate(
        inicio.getDate() - 6
    );

    inicio.setHours(
        0,
        0,
        0,
        0
    );

    return {
        inicio,
        fim
    };
}


function obterSemanaDaData(data) {
    const inicio =
        new Date(data);

    const diaSemana =
        inicio.getDay();

    const diferenca =
        diaSemana === 0
            ? -6
            : 1 - diaSemana;

    inicio.setDate(
        inicio.getDate() + diferenca
    );

    inicio.setHours(
        0,
        0,
        0,
        0
    );

    const fim =
        new Date(inicio);

    fim.setDate(
        fim.getDate() + 6
    );

    fim.setHours(
        23,
        59,
        59,
        999
    );

    return {
        inicio,
        fim
    };
}


function filtrarPorIntervalo(
    base,
    inicio,
    fim
) {
    return base.filter(linha => {
        const data =
            obterDataOcorrencia(linha);

        return (
            data &&
            data >= inicio &&
            data <= fim
        );
    });
}


function atualizarTextoPeriodo(contexto) {
    if (!periodoSelecionadoTexto) return;

    if (periodoSelecionado === "ano") {
        periodoSelecionadoTexto.textContent =
            `Ano selecionado: ` +
            `${contexto.periodoAtual.inicio.getFullYear()}`;

        return;
    }

    if (periodoSelecionado === "mes") {
        periodoSelecionadoTexto.textContent =
            `Mês selecionado: ` +
            formatarMesAno(
                contexto.periodoAtual.inicio
            );

        return;
    }

    periodoSelecionadoTexto.textContent =
        `Semana selecionada: ` +
        `${formatarDataCurta(
            contexto.periodoAtual.inicio
        )} a ${formatarDataCurta(
            contexto.periodoAtual.fim
        )}`;
}


/* =========================================================
   BATERIAS — BAHIA E GOIÁS
========================================================= */

function atualizarBlocoBaterias(contexto) {
    const bateriasAtual =
        filtrarBateriasBahiaGoias(
            contexto.linhasAtual
        );

    const bateriasAnterior =
        filtrarBateriasBahiaGoias(
            contexto.linhasAnterior
        );

    const totalAtual =
        somarQuantidade(bateriasAtual);

    const totalAnterior =
        somarQuantidade(
            bateriasAnterior
        );

    const valorAtual =
        somarValor(bateriasAtual);

    const valorAnterior =
        somarValor(bateriasAnterior);

    const ocorrenciasAtual =
        contarIdsUnicos(bateriasAtual);

    const ocorrenciasAnterior =
        contarIdsUnicos(
            bateriasAnterior
        );    

    alterarTexto(
        "bateriasTotal",
        formatarNumero(totalAtual)
    );

    alterarTexto(
        "bateriasMediaSemanal",
        formatarNumero(
            calcularMediaSemanal(
                bateriasAtual,
                contexto.periodoAtual
            )
        )
    );

    alterarTexto(
        "bateriasValorTotal",
        formatarMoeda(valorAtual)
    );

    alterarTexto(
        "bateriasOcorrencias",
        formatarNumero(
            ocorrenciasAtual
        )
    );

    atualizarVariacao(
        "bateriasVariacao",
        totalAtual,
        totalAnterior
    );

    atualizarVariacao(
        "bateriasValorVariacao",
        valorAtual,
        valorAnterior
    );

    atualizarVariacao(
        "bateriasOcorrenciasVariacao",
        ocorrenciasAtual,
        ocorrenciasAnterior
    );

    alterarTexto(
        "bateriasPeriodoAnterior",
        `Período anterior: ` +
        `${formatarNumero(totalAnterior)}`
    );

    alterarTexto(
        "bateriasValorAnterior",
        `Período anterior: ` +
        `${formatarMoeda(valorAnterior)}`
    );

    alterarTexto(
        "bateriasOcorrenciasAnterior",
        `Período anterior: ` +
        `${formatarNumero(
            ocorrenciasAnterior
        )}`
    );

    criarGraficoBateriasPorCidade(
        bateriasAtual
    );
    
    criarGraficoBateriasMensalBahia(contexto.linhasAtual);
    criarGraficoBateriasMensalGoias(baseSIGO2026);
    criarGraficoTop10SitesBahia(baseSIGO2026);
}


function filtrarBateriasBahiaGoias(
    linhas
) {
    return linhas.filter(linha => {
        const tipo =
            obterTipoPrejuizo(linha);

        const estado =
            obterEstado(linha);

        const ehBateria =
            tipo.includes("BATERIA");

        const estadoPermitido =
            estado === "BAHIA" ||
            estado === "GOIAS";

        return (
            ehBateria &&
            estadoPermitido
        );
    });
}


function criarGraficoBateriasPorCidade(
    linhas
) {
    const dados = {
        "Salvador": 0,
        "Camaçari": 0,
        "Simões Filho": 0,
        "Vera Cruz": 0,
        "Outras cidades BA": 0,
        "Goiânia": 0,
        "Outras cidades GO": 0
    };

    linhas.forEach(linha => {
        const estado =
            obterEstado(linha);

        const cidade =
            normalizarTexto(
                obterCampo(
                    linha,
                    ["Cidade"]
                )
            );

        const quantidade =
            obterQuantidade(linha);

        if (estado === "BAHIA") {
            if (cidade === "SALVADOR") {
                dados["Salvador"] +=
                    quantidade;
            } else if (
                cidade === "CAMACARI"
            ) {
                dados["Camaçari"] +=
                    quantidade;
            } else if (
                cidade === "SIMOES FILHO"
            ) {
                dados["Simões Filho"] +=
                    quantidade;
            } else if (
                cidade === "VERA CRUZ"
            ) {
                dados["Vera Cruz"] +=
                    quantidade;
            } else {
                dados["Outras cidades BA"] +=
                    quantidade;
            }
        }

        if (estado === "GOIAS") {
            if (cidade === "GOIANIA") {
                dados["Goiânia"] +=
                    quantidade;
            } else {
                dados["Outras cidades GO"] +=
                    quantidade;
            }
        }
    });

    destruirGrafico(
        graficoBateriasCidades
    );

    graficoBateriasCidades =
        new Chart(
            document.getElementById(
                "graficoBateriasCidades"
            ),
            {
                type: "bar",

                data: {
                    labels:
                        Object.keys(dados),

                    datasets: [
                        {
                            label:
                                "Quantidade de baterias",

                            data:
                                Object.values(dados),

                            backgroundColor:
                                "#ef233c",

                            borderColor:
                                "#ff4d5f",

                            borderWidth: 1,

                            borderRadius: 24,

                            borderSkipped:
                                false,

                            maxBarThickness: 65
                        }
                    ]
                },

                options:
                    opcoesGraficoBarras()
            }
        );
}


/* =========================================================
   CABOS — ESTADO DE SÃO PAULO
========================================================= */

function atualizarBlocoCabos(contexto) {
    const cabosAtual =
        filtrarCabosSaoPaulo(
            contexto.linhasAtual
        );

    const cabosAnterior =
        filtrarCabosSaoPaulo(
            contexto.linhasAnterior
        );

    const totalAtual =
        somarQuantidade(cabosAtual);

    const totalAnterior =
        somarQuantidade(cabosAnterior);

    const valorAtual =
        somarValor(cabosAtual);

    const valorAnterior =
        somarValor(cabosAnterior);

    const ocorrenciasAtual =
        contarIdsUnicos(cabosAtual);

    const ocorrenciasAnterior =
        contarIdsUnicos(cabosAnterior);

    alterarTexto(
        "cabosTotal",
        formatarNumero(totalAtual)
    );

    alterarTexto(
        "cabosMediaSemanal",
        formatarNumero(
            calcularMediaSemanal(
                cabosAtual,
                contexto.periodoAtual
            )
        )
    );

    alterarTexto(
        "cabosValorTotal",
        formatarMoeda(valorAtual)
    );

    alterarTexto(
        "cabosOcorrencias",
        formatarNumero(
            ocorrenciasAtual
        )
    );

    atualizarVariacao(
        "cabosVariacao",
        totalAtual,
        totalAnterior
    );

    atualizarVariacao(
        "cabosValorVariacao",
        valorAtual,
        valorAnterior
    );

    atualizarVariacao(
        "cabosOcorrenciasVariacao",
        ocorrenciasAtual,
        ocorrenciasAnterior
    );

    alterarTexto(
        "cabosPeriodoAnterior",
        `Período anterior: ` +
        `${formatarNumero(totalAnterior)}`
    );

    alterarTexto(
        "cabosValorAnterior",
        `Período anterior: ` +
        `${formatarMoeda(valorAnterior)}`
    );

    alterarTexto(
        "cabosOcorrenciasAnterior",
        `Período anterior: ` +
        `${formatarNumero(
            ocorrenciasAnterior
        )}`
    );

    criarGraficoCabosPorCidade(
        cabosAtual
    );
}


function filtrarCabosSaoPaulo(
    linhas
) {
    return linhas.filter(linha => {
        const tipo =
            obterTipoPrejuizo(linha);

        const estado =
            obterEstado(linha);

        const ehCabo =
            tipo === "CABO" ||
            tipo === "CABOS" ||
            tipo.includes("CABO");

        return (
            ehCabo &&
            estado === "SAO PAULO"
        );
    });
}


function criarGraficoCabosPorCidade(
    linhas
) {
    const dados = {
        "Guarulhos": 0,
        "Outras cidades SP": 0
    };

    linhas.forEach(linha => {
        const cidade =
            normalizarTexto(
                obterCampo(
                    linha,
                    ["Cidade"]
                )
            );

        const quantidade =
            obterQuantidade(linha);

        if (cidade === "GUARULHOS") {
            dados["Guarulhos"] +=
                quantidade;
        } else {
            dados["Outras cidades SP"] +=
                quantidade;
        }
    });

    destruirGrafico(
        graficoCabosCidades
    );

    graficoCabosCidades =
        new Chart(
            document.getElementById(
                "graficoCabosCidades"
            ),
            {
                type: "bar",

                data: {
                    labels:
                        Object.keys(dados),

                    datasets: [
                        {
                            label:
                                "Quantidade de cabos",

                            data:
                                Object.values(dados),

                            backgroundColor:
                                "#8b5cf6",

                            borderColor:
                                "#a879ff",

                            borderWidth: 1,

                            borderRadius: 24,

                            borderSkipped:
                                false,

                            maxBarThickness: 100
                        }
                    ]
                },

                options:
                    opcoesGraficoBarras()
            }
        );
}


/* =========================================================
   OCORRÊNCIAS POR ESTADO
========================================================= */

function atualizarOcorrenciasEstados(
    contexto
) {
    const linhasAtual =
        filtrarEstadosMonitorados(
            contexto.linhasAtual
        );

    const linhasAnterior =
        filtrarEstadosMonitorados(
            contexto.linhasAnterior
        );

    const totalAtual =
        contarIdsUnicos(linhasAtual);

    const totalAnterior =
        contarIdsUnicos(
            linhasAnterior
        );

    alterarTexto(
        "ocorrenciasTotal",
        formatarNumero(totalAtual)
    );

    atualizarVariacao(
        "ocorrenciasVariacao",
        totalAtual,
        totalAnterior
    );

    alterarTexto(
        "ocorrenciasPeriodoAnterior",
        `Período anterior: ` +
        `${formatarNumero(totalAnterior)}`
    );

    const valores =
        estadosMonitorados.map(
            estado => {
                const linhasEstado =
                    linhasAtual.filter(
                        linha =>
                            obterEstado(linha) ===
                            estado
                    );

                return contarIdsUnicos(
                    linhasEstado
                );
            }
        );

    destruirGrafico(
        graficoOcorrenciasEstados
    );

    graficoOcorrenciasEstados =
        new Chart(
            document.getElementById(
                "graficoOcorrenciasEstados"
            ),
            {
                type: "bar",

                data: {
                    labels: [
                        "Bahia",
                        "Goiás",
                        "São Paulo"
                    ],

                    datasets: [
                        {
                            label:
                                "Ocorrências",

                            data: valores,

                            backgroundColor:
                                "#22c55e",

                            borderColor:
                                "#4ade80",

                            borderWidth: 1,

                            borderRadius: 24,

                            borderSkipped:
                                false,

                            maxBarThickness: 85
                        }
                    ]
                },

                options:
                    opcoesGraficoBarras()
            }
        );
}


function filtrarEstadosMonitorados(
    linhas
) {
    return linhas.filter(linha =>
        estadosMonitorados.includes(
            obterEstado(linha)
        )
    );
}


/* =========================================================
   RESUMO FINANCEIRO
========================================================= */

function atualizarResumoFinanceiro(
    contexto
) {
    const linhasAtual =
        filtrarEstadosMonitorados(
            contexto.linhasAtual
        );

    const linhasAnterior =
        filtrarEstadosMonitorados(
            contexto.linhasAnterior
        );

    atualizarFinanceiroEstado(
        "BAHIA",
        "valorBahia",
        "valorBahiaVariacao",
        "valorBahiaAnterior",
        linhasAtual,
        linhasAnterior
    );

    atualizarFinanceiroEstado(
        "GOIAS",
        "valorGoias",
        "valorGoiasVariacao",
        "valorGoiasAnterior",
        linhasAtual,
        linhasAnterior
    );

    atualizarFinanceiroEstado(
        "SAO PAULO",
        "valorSaoPaulo",
        "valorSaoPauloVariacao",
        "valorSaoPauloAnterior",
        linhasAtual,
        linhasAnterior
    );

    const totalAtual =
        somarValor(linhasAtual);

    const totalAnterior =
        somarValor(linhasAnterior);

    alterarTexto(
        "resumoValor",
        formatarMoeda(totalAtual)
    );

    atualizarVariacao(
        "resumoValorVariacao",
        totalAtual,
        totalAnterior
    );

    alterarTexto(
        "resumoValorAnterior",
        `Período anterior: ` +
        `${formatarMoeda(totalAnterior)}`
    );
}


function atualizarFinanceiroEstado(
    estado,
    idValor,
    idVariacao,
    idAnterior,
    linhasAtual,
    linhasAnterior
) {
    const atual =
        somarValor(
            linhasAtual.filter(
                linha =>
                    obterEstado(linha) ===
                    estado
            )
        );

    const anterior =
        somarValor(
            linhasAnterior.filter(
                linha =>
                    obterEstado(linha) ===
                    estado
            )
        );

    alterarTexto(
        idValor,
        formatarMoeda(atual)
    );

    atualizarVariacao(
        idVariacao,
        atual,
        anterior
    );

    alterarTexto(
        idAnterior,
        `Período anterior: ` +
        `${formatarMoeda(anterior)}`
    );
}


/* =========================================================
   CONFIGURAÇÃO DOS GRÁFICOS
========================================================= */

function opcoesGraficoBarras() {
    return {
        responsive: true,
        maintainAspectRatio: false,

        layout: {
            padding: {
                top: 30,
                right: 20
            }
        },

        plugins: {
            legend: {
                display: false
            },

            tooltip: {
                backgroundColor:
                    "#0f172a",

                titleColor:
                    "#ffffff",

                bodyColor:
                    "#ffffff",

                padding: 12,
                cornerRadius: 14
            }
        },

        scales: {
            x: {
                ticks: {
                    color: "#d1d5db",
                    font: {
                        size: 12
                    }
                },

                grid: {
                    display: false
                },

                border: {
                    color:
                        "rgba(148, 163, 184, 0.35)"
                }
            },

            y: {
                beginAtZero: true,

                ticks: {
                    color: "#9ca3af"
                },

                grid: {
                    color:
                        "rgba(148, 163, 184, 0.12)"
                },

                border: {
                    display: false
                }
            }
        }
    };
}


function destruirGrafico(grafico) {
    if (grafico) {
        grafico.destroy();
    }
}


/* =========================================================
   CÁLCULOS
========================================================= */

function somarQuantidade(linhas) {
    return linhas.reduce(
        (total, linha) =>
            total + obterQuantidade(linha),
        0
    );
}


function somarValor(linhas) {
    return linhas.reduce(
        (total, linha) =>
            total + obterValor(linha),
        0
    );
}


function obterQuantidade(linha) {
    return converterNumero(
        obterCampo(
            linha,
            ["Quantidade"]
        )
    );
}


function obterValor(linha) {
    return converterNumero(
        obterCampo(
            linha,
            [
                "Valor Total",
                "Valor"
            ]
        )
    );
}


function obterEstado(linha) {
    return normalizarTexto(
        obterCampo(linha, ["Estado"])
    );
}


function obterTipoPrejuizo(linha) {
    return normalizarTexto(
        obterCampo(
            linha,
            [
                "Item de Prejuízo / Recuperação",
                "Tipo de Prejuízo / Recuperação",
                "Tipos de Prejuízo / Recuperação"
            ]
        )
    );
}


function contarIdsUnicos(linhas) {
    const ids = new Set();

    linhas.forEach(linha => {
        const id =
            limparTexto(
                obterCampo(
                    linha,
                    [
                        "Id Ocorrência",
                        "ID Ocorrência",
                        "Id"
                    ]
                )
            );

        if (id !== "") {
            ids.add(id);
        }
    });

    return ids.size;
}


function calcularMediaSemanal(
    linhas,
    intervalo
) {
    const milissegundosDia =
        1000 * 60 * 60 * 24;

    const dias =
        Math.max(
            1,
            Math.round(
                (
                    intervalo.fim -
                    intervalo.inicio
                ) /
                milissegundosDia
            ) + 1
        );

    const semanas =
        Math.max(
            1,
            dias / 7
        );

    return Math.round(
        somarQuantidade(linhas) /
        semanas
    );
}


function calcularVariacao(
    atual,
    anterior
) {
    if (anterior === 0) {
        return atual === 0
            ? 0
            : 100;
    }

    return (
        (
            atual - anterior
        ) /
        anterior
    ) * 100;
}


/* =========================================================
   VARIAÇÃO
   AUMENTO = VERMELHO
   REDUÇÃO = VERDE
========================================================= */

function atualizarVariacao(
    id,
    atual,
    anterior
) {
    const elemento =
        document.getElementById(id);

    if (!elemento) return;

    const percentual =
        calcularVariacao(
            atual,
            anterior
        );

    elemento.classList.remove(
        "variacao-positiva",
        "variacao-negativa"
    );

    if (percentual > 0) {
        elemento.textContent =
            `▲ ${formatarPercentual(
                percentual
            )}`;

        elemento.classList.add(
            "variacao-positiva"
        );

        return;
    }

    if (percentual < 0) {
        elemento.textContent =
            `▼ ${formatarPercentual(
                Math.abs(percentual)
            )}`;

        elemento.classList.add(
            "variacao-negativa"
        );

        return;
    }

    elemento.textContent = "— 0,0%";
}


/* =========================================================
   DATAS
========================================================= */

function obterDataOcorrencia(linha) {
    const valor =
        obterCampo(
            linha,
            [
                "Data da Ocorrência",
                "Data da Ocorrencia",
                "DATA DA OCORRÊNCIA",
                "DATA DA OCORRENCIA"
            ]
        );

    return converterData(valor);
}


function converterData(valor) {
    if (!valor) return null;

    if (
        valor instanceof Date &&
        !Number.isNaN(valor.getTime())
    ) {
        return new Date(valor);
    }

    if (typeof valor === "number") {
        const dataExcel =
            XLSX.SSF.parse_date_code(
                valor
            );

        if (dataExcel) {
            return new Date(
                dataExcel.y,
                dataExcel.m - 1,
                dataExcel.d,
                dataExcel.H || 0,
                dataExcel.M || 0,
                Math.floor(
                    dataExcel.S || 0
                )
            );
        }
    }

    const texto =
        limparTexto(valor);

    const dataBrasileira =
        texto.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
        );

    if (dataBrasileira) {
        return new Date(
            Number(dataBrasileira[3]),
            Number(dataBrasileira[2]) - 1,
            Number(dataBrasileira[1])
        );
    }

    const dataIso =
        new Date(texto);

    if (
        !Number.isNaN(
            dataIso.getTime()
        )
    ) {
        return dataIso;
    }

    return null;
}


function encontrarDataReferencia(base) {
    const datas = base
        .map(obterDataOcorrencia)
        .filter(data =>
            data &&
            !Number.isNaN(
                data.getTime()
            )
        );

    if (datas.length === 0) {
        return new Date();
    }

    return new Date(
        Math.max(
            ...datas.map(data =>
                data.getTime()
            )
        )
    );
}


/* =========================================================
   ACESSO AOS CAMPOS
========================================================= */

function obterCampo(
    linha,
    nomesPossiveis
) {
    const chaves =
        Object.keys(linha);

    for (
        const nomePossivel
        of nomesPossiveis
    ) {
        const nomeNormalizado =
            normalizarCabecalho(
                nomePossivel
            );

        const chaveEncontrada =
            chaves.find(chave =>
                normalizarCabecalho(
                    chave
                ) === nomeNormalizado
            );

        if (chaveEncontrada) {
            return linha[
                chaveEncontrada
            ];
        }
    }

    return "";
}


/* =========================================================
   CONVERSÃO E FORMATAÇÃO
========================================================= */

function converterNumero(valor) {
    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return 0;
    }

    if (typeof valor === "number") {
        return Number.isFinite(valor)
            ? valor
            : 0;
    }

    let texto =
        String(valor)
            .replace(/\s/g, "")
            .replace(/R\$/gi, "")
            .trim();

    if (
        texto.includes(".") &&
        texto.includes(",")
    ) {
        texto = texto
            .replace(/\./g, "")
            .replace(",", ".");
    } else if (
        texto.includes(",")
    ) {
        texto =
            texto.replace(",", ".");
    }

    const numero =
        Number(texto);

    return Number.isFinite(numero)
        ? numero
        : 0;
}


function formatarNumero(valor) {
    return Number(valor || 0)
        .toLocaleString(
            "pt-BR",
            {
                maximumFractionDigits: 0
            }
        );
}


function formatarMoeda(valor) {
    return Number(valor || 0)
        .toLocaleString(
            "pt-BR",
            {
                style: "currency",
                currency: "BRL"
            }
        );
}


function formatarPercentual(valor) {
    return (
        Number(valor || 0)
            .toLocaleString(
                "pt-BR",
                {
                    minimumFractionDigits: 1,
                    maximumFractionDigits: 1
                }
            ) +
        "%"
    );
}


function formatarDataCurta(data) {
    return data.toLocaleDateString(
        "pt-BR"
    );
}


function formatarMesAno(data) {
    const texto =
        data.toLocaleDateString(
            "pt-BR",
            {
                month: "long",
                year: "numeric"
            }
        );

    return texto.charAt(0)
        .toUpperCase() +
        texto.slice(1);
}


function alterarTexto(id, valor) {
    const elemento =
        document.getElementById(id);

    if (elemento) {
        elemento.textContent = valor;
    }
}


/* =========================================================
   NORMALIZAÇÃO
========================================================= */

function normalizarCabecalho(texto) {
    let normalizado =
        limparTexto(texto)
            .normalize("NFD")
            .replace(
                /[\u0300-\u036f]/g,
                ""
            )
            .replace(/\s+/g, " ")
            .toLowerCase();

    if (
        aliasesCabecalhos[
            normalizado
        ]
    ) {
        normalizado =
            aliasesCabecalhos[
                normalizado
            ];
    }

    return normalizado;
}


function normalizarTexto(texto) {
    return limparTexto(texto)
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(/\s+/g, " ")
        .toUpperCase();
}


function limparTexto(texto) {
    return String(texto ?? "")
        .replace(/\u00a0/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

let graficoBateriasMensalBahia = null;

function criarGraficoBateriasMensalBahia(linhas){

    const meses = [
        "Jan","Fev","Mar","Abr","Mai","Jun",
        "Jul","Ago","Set","Out","Nov","Dez"
    ];

    const valores = new Array(12).fill(0);

    linhas.forEach(linha=>{

        const estado = obterEstado(linha);

        if (estado !== "BAHIA") return;

        const tipo = obterTipoPrejuizo(linha);

        if(!tipo.includes("BATERIA")) return;

        const data = obterDataOcorrencia(linha);

        if(!data) return;

        valores[data.getMonth()] += obterQuantidade(linha);

    });

    if(graficoBateriasMensalBahia){
        graficoBateriasMensalBahia.destroy();
    }

    graficoBateriasMensalBahia = new Chart(

        document.getElementById("graficoBateriasMensalBahia"),

        {

            type:"bar",

            data:{

                labels:meses,

                datasets: [{
                    data: valores,

                    backgroundColor: "#f35810",
                    hoverBackgroundColor: "#d1d5db",

                    borderColor: "#f35810",
                    borderWidth: 1,

                    hoverBorderColor: "#ffffff",
                    hoverBorderWidth: 2,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 45
                }]

            },

            options:opcoesGraficoBarras()

        }

    );

}

function criarGraficoBateriasMensalGoias(linhas) {
    let labels = [];
    let valores = [];

    const linhasGoias = linhas.filter(linha => {
        const estado = obterEstado(linha);
        const tipo = obterTipoPrejuizo(linha);

        return (
            estado === "GOIAS" &&
            tipo.includes("BATERIA")
        );
    });

    if (periodoSelecionado === "ano") {
        labels = [
            "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"
        ];

        valores = new Array(12).fill(0);

        const anoReferencia =
            dataReferencia.getFullYear();

        linhasGoias.forEach(linha => {
            const data = obterDataOcorrencia(linha);

            if (!data) return;
            if (data.getFullYear() !== anoReferencia) return;

            valores[data.getMonth()] +=
                obterQuantidade(linha);
        });
    }

    if (periodoSelecionado === "mes") {
        const contexto =
            criarContextoPeriodo(baseSIGO2026);

        const dadosMes =
            agruparPorSemanasDoMes(
                linhasGoias,
                contexto.referencia
            );

        labels = dadosMes.labels;
        valores = dadosMes.valores;
    }

    if (periodoSelecionado === "semana") {
        const contexto =
            criarContextoPeriodo(baseSIGO2026);

        const diasSemana = [
            "Seg", "Ter", "Qua",
            "Qui", "Sex", "Sáb", "Dom"
        ];

        labels = diasSemana;
        valores = new Array(7).fill(0);

        const linhasSemana =
            filtrarPorIntervalo(
                linhasGoias,
                contexto.periodoAtual.inicio,
                contexto.periodoAtual.fim
            );

        linhasSemana.forEach(linha => {
            const data =
                obterDataOcorrencia(linha);

            if (!data) return;

            const dia =
                data.getDay() === 0
                    ? 6
                    : data.getDay() - 1;

            valores[dia] +=
                obterQuantidade(linha);
        });
    }

    if (graficoBateriasMensalGoias) {
        graficoBateriasMensalGoias.destroy();
    }

    graficoBateriasMensalGoias = new Chart(
        document.getElementById(
            "graficoBateriasMensalGoias"
        ),
        {
            type: "bar",

            data: {
                labels,

                datasets: [{
                    data: valores,

                    backgroundColor: "#c5b916",
                    hoverBackgroundColor: "#d1d5db",

                    borderColor: "#c5b916",
                    borderWidth: 1,

                    hoverBorderColor: "#ffffff",
                    hoverBorderWidth: 2,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 45
                }]
            },

            options: opcoesGraficoBarras()
        }
    );
}

function criarGraficoTop10SitesBahia(linhas) {
    const canvas = document.getElementById(
        "graficoTop10SitesBahia"
    );

    if (!canvas) {
        console.error(
            "Canvas graficoTop10SitesBahia não encontrado no HTML."
        );
        return;
    }

    const totaisPorSite = {};

    linhas.forEach(linha => {
        const estado = obterEstado(linha);
        const tipo = obterTipoPrejuizo(linha);

        const site = limparTexto(
            obterCampo(
                linha,
                [
                    "Site/Loja",
                    "Site / Loja"
                ]
            )
        );

        const quantidade =
            obterQuantidade(linha);

        if (estado !== "BAHIA") return;
        if (!tipo.includes("BATERIA")) return;
        if (site === "") return;

        totaisPorSite[site] =
            (totaisPorSite[site] || 0) +
            quantidade;
    });

    const top10 = Object.entries(totaisPorSite)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const labels =
        top10.map(item => item[0]);

    const valores =
        top10.map(item => item[1]);

    if (graficoTop10SitesBahia) {
        graficoTop10SitesBahia.destroy();
    }

    graficoTop10SitesBahia = new Chart(
        canvas,
        {
            type: "bar",

            data: {
                labels,

                datasets: [{
                    label: "Quantidade de baterias",
                    data: valores,

                    backgroundColor: "#f35810",
                    hoverBackgroundColor: "#d1d5db",

                    borderColor: "#f35810",
                    borderWidth: 1,

                    hoverBorderColor: "#ffffff",
                    hoverBorderWidth: 2,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 35
                }]
            },

            options: {
                ...opcoesGraficoBarras(),

                indexAxis: "y",

                scales: {
                    x: {
                        beginAtZero: true,

                        ticks: {
                            color: "#9ca3af",
                            precision: 0
                        },

                        grid: {
                            color:
                                "rgba(148, 163, 184, 0.12)"
                        }
                    },

                    y: {
                        ticks: {
                            color: "#d1d5db"
                        },

                        grid: {
                            display: false
                        }
                    }
                }
            }
        }
    );
}

/* =========================================================
   ANIMAÇÕES AO ROLAR A PÁGINA
========================================================= */

function configurarAnimacoesScroll() {
    const cards = document.querySelectorAll(
        ".indicador-card, .resumo-card, .dashboard-card"
    );

    const graficos = document.querySelectorAll(
        ".grafico-painel"
    );

    const blocos = document.querySelectorAll(
        ".importacao-sigo, .periodo-analise, .titulo-bloco"
    );

    cards.forEach((card, indice) => {
        card.classList.add("animar-scroll");

        if (indice % 2 === 0) {
            card.classList.add("entrar-esquerda");
        } else {
            card.classList.add("entrar-direita");
        }
    });

    graficos.forEach((grafico, indice) => {
        grafico.classList.add(
            "animar-scroll",
            indice % 2 === 0
                ? "entrar-esquerda"
                : "entrar-direita"
        );
    });

    blocos.forEach(bloco => {
        bloco.classList.add(
            "animar-scroll",
            "entrar-baixo"
        );
    });

    const observador = new IntersectionObserver(
        entradas => {
            entradas.forEach(entrada => {
                if (!entrada.isIntersecting) return;

                const elemento = entrada.target;

                elemento.classList.add("visivel");

                const canvas =
                    elemento.querySelector("canvas");

                if (canvas) {
                    const grafico =
                        Chart.getChart(canvas);

                    if (grafico) {
                        grafico.reset();

                        setTimeout(() => {
                            grafico.update();
                        }, 180);
                    }
                }

                observador.unobserve(elemento);
            });
        },
        {
            threshold: 0.15,
            rootMargin: "0px 0px -40px 0px"
        }
    );

    document
        .querySelectorAll(".animar-scroll")
        .forEach(elemento => {
            observador.observe(elemento);
        });
}