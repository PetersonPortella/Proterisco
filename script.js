/* =========================================================
   NAVEGAÇÃO
========================================================= */

function mostrarSecao(id) {
    document.querySelectorAll(".secao").forEach(secao => {
        secao.classList.remove("ativa");
    });

    const secao = document.getElementById(id);
    if (secao) secao.classList.add("ativa");
}

/* =========================================================
   CONFIGURAÇÕES
========================================================= */

const caminhoPlanilhaMestra = "./dados/SIGO_2026.xlsx";
const estadosMonitorados = ["BAHIA", "GOIAS", "SAO PAULO"];

const aliasesCabecalhos = {
    "tipos de prejuizo / recuperacao": "item de prejuizo / recuperacao",
    "tipo de prejuizo / recuperacao": "item de prejuizo / recuperacao",
    "valor": "valor total"
};

let workbookAtualizado = null;
let cabecalhosMestra = [];
let linhasMestra = [];
let linhasNovas = [];
let baseSIGO2026 = [];
let baseProcessada = [];

let periodoSelecionado = "semana";
let dataReferencia = new Date();

const graficos = Object.create(null);

const mesesCurtos = [
    "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
    "Jul", "Ago", "Set", "Out", "Nov", "Dez"
];

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
   PLUGIN DE VALORES NOS GRÁFICOS
========================================================= */

const pluginValores = {
    id: "pluginValores",

    afterDatasetsDraw(chart) {
        if (window.innerWidth <= 600) return;

        const contexto = chart.ctx;

        contexto.save();
        contexto.fillStyle = "#f8fafc";
        contexto.font = "bold 12px Arial";

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
                                posicao.y - 7
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
    document
        .querySelectorAll(".botao-periodo")
        .forEach(botao => {
            botao.addEventListener(
                "click",
                () => {
                    document
                        .querySelectorAll(".botao-periodo")
                        .forEach(item => {
                            item.classList.remove("ativo");
                        });

                    botao.classList.add("ativo");

                    periodoSelecionado =
                        botao.dataset.periodo || "semana";

                    requestAnimationFrame(
                        atualizarDashboard
                    );
                }
            );
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

        const leitura =
            lerDadosDaAba(abaMestra);

        cabecalhosMestra =
            leitura.cabecalhos;

        linhasMestra =
            leitura.linhas;

        baseSIGO2026 = [
            ...linhasMestra
        ];

        reconstruirBaseProcessada();

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
   IMPORTAÇÃO DA PLANILHA SEMANAL
========================================================= */

inputExcel.addEventListener(
    "change",
    async evento => {
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

            if (!linhasNovas.length) {
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
            reconstruirBaseProcessada();

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
   LEITURA DE XLS/XLSX
========================================================= */

async function lerPlanilhaSemanal(arquivo) {
    const conteudo =
        await arquivo.arrayBuffer();

    if (
        arquivo.name
            .toLowerCase()
            .endsWith(".xls") &&
        arquivoEhHtml(conteudo)
    ) {
        return lerArquivoHtmlXls(
            conteudo
        );
    }

    const workbook = XLSX.read(
        conteudo,
        {
            type: "array",
            cellDates: true
        }
    );

    const nomeAba =
        workbook.SheetNames[0];

    return lerDadosDaAba(
        workbook.Sheets[nomeAba]
    ).linhas;
}

function lerDadosDaAba(aba) {
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
        dados[indiceCabecalho]
            .map(valor =>
                limparTexto(valor)
            );

    const linhas = dados
        .slice(indiceCabecalho + 1)
        .filter(linha =>
            linha.some(valor =>
                limparTexto(valor) !== ""
            )
        )
        .map(linha =>
            transformarLinhaEmObjeto(
                linha,
                cabecalhos
            )
        );

    return {
        cabecalhos:
            cabecalhos.filter(Boolean),

        linhas
    };
}

function arquivoEhHtml(conteudo) {
    const bytes =
        new Uint8Array(
            conteudo.slice(0, 300)
        );

    const texto =
        new TextDecoder("windows-1252")
            .decode(bytes)
            .trim()
            .toLowerCase();

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
        new DOMParser()
            .parseFromString(
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
            limparTexto(
                celula.textContent
            )
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
                    if (cabecalho) {
                        objeto[cabecalho] =
                            limparTexto(
                                celulas[indice]
                                    ?.textContent ?? ""
                            );
                    }
                }
            );

            return objeto;
        })
        .filter(linha =>
            Object.values(linha)
                .some(valor =>
                    limparTexto(valor) !== ""
                )
        );
}

function localizarCabecalhoTabelaHtml(
    linhas
) {
    return linhas.findIndex(
        linha => {
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

            const encontrados = [
                "estado",
                "grupo de ocorrencia",
                "quantidade"
            ].filter(cabecalho =>
                valores.includes(cabecalho)
            );

            return encontrados.length >= 2;
        }
    );
}

function localizarLinhaCabecalho(dados) {
    const esperados = [
        "id ocorrencia",
        "estado",
        "grupo de ocorrencia",
        "quantidade"
    ];

    return dados.findIndex(
        linha => {
            const valores =
                linha.map(
                    normalizarCabecalho
                );

            return esperados
                .filter(cabecalho =>
                    valores.includes(cabecalho)
                )
                .length >= 3;
        }
    );
}

function transformarLinhaEmObjeto(
    linha,
    cabecalhos
) {
    const objeto = {};

    cabecalhos.forEach(
        (cabecalho, indice) => {
            if (cabecalho) {
                objeto[cabecalho] =
                    linha[indice] ?? "";
            }
        }
    );

    return objeto;
}

/* =========================================================
   VALIDAÇÃO E PADRONIZAÇÃO
========================================================= */

function validarCabecalhos(
    cabecalhosDaMestra,
    novasLinhas
) {
    const cabecalhosNova =
        new Set(
            Object.keys(novasLinhas[0])
                .map(normalizarCabecalho)
        );

    const faltando =
        cabecalhosDaMestra.filter(
            cabecalho => {
                return !cabecalhosNova.has(
                    normalizarCabecalho(
                        cabecalho
                    )
                );
            }
        );

    if (faltando.length) {
        throw new Error(
            `Planilha inválida. Colunas ausentes: ` +
            `${faltando.join(", ")}`
        );
    }
}

function padronizarLinhasNovas(
    novasLinhas,
    cabecalhosDaMestra
) {
    return novasLinhas.map(
        linha => {
            const mapa =
                Object.create(null);

            Object.keys(linha)
                .forEach(cabecalho => {
                    mapa[
                        normalizarCabecalho(
                            cabecalho
                        )
                    ] = linha[cabecalho];
                });

            const padronizada = {};

            cabecalhosDaMestra
                .forEach(cabecalho => {
                    padronizada[cabecalho] =
                        mapa[
                            normalizarCabecalho(
                                cabecalho
                            )
                        ] ?? "";
                });

            return padronizada;
        }
    );
}

/* =========================================================
   SALVAR PLANILHA ATUALIZADA
========================================================= */

function atualizarWorkbook() {
    const nomeAba =
        workbookAtualizado.SheetNames[0];

    const dados = [
        cabecalhosMestra,

        ...baseSIGO2026.map(linha =>
            cabecalhosMestra.map(
                cabecalho =>
                    linha[cabecalho] ?? ""
            )
        )
    ];

    workbookAtualizado
        .Sheets[nomeAba] =
        XLSX.utils.aoa_to_sheet(
            dados
        );
}

salvarPlanilha.addEventListener(
    "click",
    () => {
        if (!workbookAtualizado) return;

        XLSX.writeFile(
            workbookAtualizado,
            "SIGO_2026_ATUALIZADA.xlsx"
        );
    }
);

/* =========================================================
   BASE PROCESSADA — UMA ÚNICA VEZ
========================================================= */

function reconstruirBaseProcessada() {
    baseProcessada =
        baseSIGO2026.map(linha => ({
            original: linha,

            data: converterData(
                obterCampoDireto(
                    linha,
                    [
                        "Data da Ocorrência",
                        "Data da Ocorrencia"
                    ]
                )
            ),

            estado: normalizarTexto(
                obterCampoDireto(
                    linha,
                    ["Estado"]
                )
            ),

            cidade: normalizarTexto(
                obterCampoDireto(
                    linha,
                    ["Cidade"]
                )
            ),

            bairro: normalizarTexto(
                obterCampoDireto(
                    linha,
                    ["Bairro"]
                )
            ),

            tipo: normalizarTexto(
                obterCampoDireto(
                    linha,
                    [
                        "Item de Prejuízo / Recuperação",
                        "Tipo de Prejuízo / Recuperação",
                        "Tipos de Prejuízo / Recuperação"
                    ]
                )
            ),

            quantidade: converterNumero(
                obterCampoDireto(
                    linha,
                    ["Quantidade"]
                )
            ),

            valor: converterNumero(
                obterCampoDireto(
                    linha,
                    [
                        "Valor Total",
                        "Valor"
                    ]
                )
            ),

            id: limparTexto(
                obterCampoDireto(
                    linha,
                    [
                        "Id Ocorrência",
                        "ID Ocorrência",
                        "Id"
                    ]
                )
            ),

            site: limparTexto(
                obterCampoDireto(
                    linha,
                    [
                        "Site/Loja",
                        "Site / Loja"
                    ]
                )
            )
        }));

    const datasValidas =
        baseProcessada
            .map(item => item.data)
            .filter(data =>
                data &&
                !Number.isNaN(
                    data.getTime()
                )
            );

    dataReferencia =
        datasValidas.length
            ? new Date(
                Math.max(
                    ...datasValidas.map(
                        data =>
                            data.getTime()
                    )
                )
            )
            : new Date();
}

function obterCampoDireto(
    linha,
    nomesPossiveis
) {
    const mapa =
        Object.create(null);

    Object.keys(linha)
        .forEach(chave => {
            mapa[
                normalizarCabecalho(chave)
            ] = linha[chave];
        });

    for (
        const nome
        of nomesPossiveis
    ) {
        const chave =
            normalizarCabecalho(nome);

        if (
            Object.prototype
                .hasOwnProperty.call(
                    mapa,
                    chave
                )
        ) {
            return mapa[chave];
        }
    }

    return "";
}

/* =========================================================
   PERÍODOS
========================================================= */

function criarContextoPeriodo() {
    const periodoAtual =
        obterIntervaloAtual(
            periodoSelecionado,
            dataReferencia
        );

    const periodoAnterior =
        obterIntervaloAnterior(
            periodoSelecionado,
            periodoAtual
        );

    return {
        periodoAtual,
        periodoAnterior,

        linhasAtual:
            filtrarProcessadaPorIntervalo(
                baseProcessada,
                periodoAtual.inicio,
                periodoAtual.fim
            ),

        linhasAnterior:
            filtrarProcessadaPorIntervalo(
                baseProcessada,
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

    return obterSemanaDaData(
        referencia
    );
}

function obterIntervaloAnterior(
    tipo,
    atual
) {
    if (tipo === "ano") {
        const ano =
            atual.inicio
                .getFullYear() - 1;

        return {
            inicio: new Date(
                ano,
                0,
                1
            ),

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
                atual.inicio
                    .getFullYear(),

                atual.inicio
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
        new Date(atual.inicio);

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

    const dia =
        inicio.getDay();

    inicio.setDate(
        inicio.getDate() +
        (
            dia === 0
                ? -6
                : 1 - dia
        )
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

function filtrarProcessadaPorIntervalo(
    base,
    inicio,
    fim
) {
    return base.filter(
        item =>
            item.data &&
            item.data >= inicio &&
            item.data <= fim
    );
}

function atualizarTextoPeriodo(
    contexto
) {
    if (!periodoSelecionadoTexto) {
        return;
    }

    if (periodoSelecionado === "ano") {
        periodoSelecionadoTexto
            .textContent =
            `Ano selecionado: ` +
            `${contexto.periodoAtual.inicio.getFullYear()}`;

        return;
    }

    if (periodoSelecionado === "mes") {
        periodoSelecionadoTexto
            .textContent =
            `Mês selecionado: ` +
            `${formatarMesAno(
                contexto.periodoAtual.inicio
            )}`;

        return;
    }

    periodoSelecionadoTexto
        .textContent =
        `Semana selecionada: ` +
        `${formatarDataCurta(
            contexto.periodoAtual.inicio
        )} a ${formatarDataCurta(
            contexto.periodoAtual.fim
        )}`;
}

/* =========================================================
   ATUALIZAR DASHBOARD
========================================================= */

function atualizarDashboard() {
    if (!baseProcessada.length) return;

    const contexto =
        criarContextoPeriodo();

    atualizarTextoPeriodo(contexto);
    atualizarBlocoBaterias(contexto);
    atualizarTop10SitesBahia(contexto);
    atualizarBlocoCabos(contexto);
    atualizarOcorrenciasEstados(contexto);
    atualizarResumoFinanceiro(contexto);
}

/* =========================================================
   BATERIAS
========================================================= */

function ehBateria(item) {
    return item.tipo.includes(
        "BATERIA"
    );
}

function atualizarBlocoBaterias(
    contexto
) {
    const bateriasPeriodoTodos =
        contexto.linhasAtual
            .filter(ehBateria);

    const atual =
        contexto.linhasAtual.filter(
            item =>
                ehBateria(item) &&
                (
                    item.estado === "BAHIA" ||
                    item.estado === "GOIAS"
                )
        );

    const anterior =
        contexto.linhasAnterior.filter(
            item =>
                ehBateria(item) &&
                (
                    item.estado === "BAHIA" ||
                    item.estado === "GOIAS"
                )
        );

    const totalAtual =
        somarQuantidadeProcessada(
            atual
        );

    const totalAnterior =
        somarQuantidadeProcessada(
            anterior
        );

    const valorAtual =
        somarValorProcessada(
            atual
        );

    const valorAnterior =
        somarValorProcessada(
            anterior
        );

    const ocorrenciasAtual =
        contarIdsProcessados(
            atual
        );

    const ocorrenciasAnterior =
        contarIdsProcessados(
            anterior
        );

    alterarTexto(
        "bateriasTotal",
        formatarNumero(totalAtual)
    );

    alterarTexto(
        "bateriasMediaSemanal",
        formatarNumero(
            calcularMediaSemanalProcessada(
                atual,
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

    atualizarGraficoBateriasCidades(
        atual
    );

    atualizarGraficoPeriodoBaterias(
        "BAHIA",
        "graficoBateriasMensalBahia",
        "#f35810"
    );

    atualizarGraficoPeriodoBaterias(
        "GOIAS",
        "graficoBateriasMensalGoias",
        "#c5b916"
    );

    atualizarGraficosBateriasPorEstadoCidade(
        bateriasPeriodoTodos
    );
}

function atualizarGraficoPeriodoBaterias(
    estado,
    canvasId,
    cor
) {
    const linhas =
        baseProcessada.filter(
            item =>
                item.estado === estado &&
                ehBateria(item)
        );

    let labels = [];
    let valores = [];

    if (periodoSelecionado === "ano") {
        labels = mesesCurtos;
        valores = new Array(12).fill(0);

        const ano =
            dataReferencia.getFullYear();

        linhas.forEach(item => {
            if (
                item.data &&
                item.data.getFullYear() === ano
            ) {
                valores[
                    item.data.getMonth()
                ] += item.quantidade;
            }
        });

    } else if (
        periodoSelecionado === "mes"
    ) {
        const agrupado =
            agruparPorSemanasDoMesProcessada(
                linhas,
                dataReferencia
            );

        labels = agrupado.labels;
        valores = agrupado.valores;

    } else {
        labels = [
            "Seg",
            "Ter",
            "Qua",
            "Qui",
            "Sex",
            "Sáb",
            "Dom"
        ];

        valores =
            new Array(7).fill(0);

        const semana =
            obterSemanaDaData(
                dataReferencia
            );

        filtrarProcessadaPorIntervalo(
            linhas,
            semana.inicio,
            semana.fim
        ).forEach(item => {
            const indice =
                item.data.getDay() === 0
                    ? 6
                    : item.data.getDay() - 1;

            valores[indice] +=
                item.quantidade;
        });
    }

    criarOuAtualizarGrafico(
        canvasId,
        {
            type: "bar",

            data: {
                labels,

                datasets: [{
                    data: valores,

                    backgroundColor: cor,
                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor: cor,
                    borderWidth: 1,

                    hoverBorderColor:
                        "#ffffff",

                    hoverBorderWidth: 2,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 45
                }]
            },

            options:
                opcoesGraficoBarras()
        }
    );
}

function agruparPorSemanasDoMesProcessada(
    linhas,
    referencia
) {
    const ano =
        referencia.getFullYear();

    const mes =
        referencia.getMonth();

    const ultimoDia =
        new Date(
            ano,
            mes + 1,
            0
        ).getDate();

    const labels = [];
    const valores = [];

    for (
        let inicioDia = 1,
            semana = 1;

        inicioDia <= ultimoDia;

        inicioDia += 7,
            semana++
    ) {
        const fimDia =
            Math.min(
                inicioDia + 6,
                ultimoDia
            );

        const inicio =
            new Date(
                ano,
                mes,
                inicioDia,
                0,
                0,
                0,
                0
            );

        const fim =
            new Date(
                ano,
                mes,
                fimDia,
                23,
                59,
                59,
                999
            );

        labels.push(
            `Semana ${semana}`
        );

        valores.push(
            somarQuantidadeProcessada(
                filtrarProcessadaPorIntervalo(
                    linhas,
                    inicio,
                    fim
                )
            )
        );
    }

    return {
        labels,
        valores
    };
}

function atualizarGraficoBateriasCidades(
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

    linhas.forEach(item => {
        if (item.estado === "BAHIA") {
            const mapa = {
                "SALVADOR": "Salvador",
                "CAMACARI": "Camaçari",
                "SIMOES FILHO":
                    "Simões Filho",
                "VERA CRUZ": "Vera Cruz"
            };

            dados[
                mapa[item.cidade] ||
                "Outras cidades BA"
            ] += item.quantidade;
        }

        if (item.estado === "GOIAS") {
            dados[
                item.cidade === "GOIANIA"
                    ? "Goiânia"
                    : "Outras cidades GO"
            ] += item.quantidade;
        }
    });

    criarOuAtualizarGrafico(
        "graficoBateriasCidades",
        {
            type: "bar",

            data: {
                labels:
                    Object.keys(dados),

                datasets: [{
                    label:
                        "Quantidade de baterias",

                    data:
                        Object.values(dados),

                    backgroundColor:
                        "#ef233c",

                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor:
                        "#ff4d5f",

                    borderWidth: 1,
                    borderRadius: 24,
                    borderSkipped: false,
                    maxBarThickness: 65
                }]
            },

            options:
                opcoesGraficoBarras()
        }
    );
}

function atualizarGraficosBateriasPorEstadoCidade(
    linhas
) {
    const bahia = {
        "Salvador": 0,
        "Camaçari": 0,
        "Simões Filho": 0,
        "Vera Cruz": 0,
        "Outras": 0
    };

    const goias = {
        "Goiânia": 0,
        "Outras": 0
    };

    const saoPaulo = {
        "Guarulhos": 0,
        "Outras": 0
    };

    linhas.forEach(item => {
        if (item.estado === "BAHIA") {
            const mapa = {
                "SALVADOR": "Salvador",
                "CAMACARI": "Camaçari",
                "SIMOES FILHO":
                    "Simões Filho",
                "VERA CRUZ": "Vera Cruz"
            };

            bahia[
                mapa[item.cidade] ||
                "Outras"
            ] += item.quantidade;
        }

        if (item.estado === "GOIAS") {
            goias[
                item.cidade === "GOIANIA"
                    ? "Goiânia"
                    : "Outras"
            ] += item.quantidade;
        }

        if (
            item.estado === "SAO PAULO"
        ) {
            saoPaulo[
                item.cidade === "GUARULHOS"
                    ? "Guarulhos"
                    : "Outras"
            ] += item.quantidade;
        }
    });

    criarGraficoOpcional(
        "graficoBateriasCidadesBahia",
        bahia,
        "#f35810"
    );

    criarGraficoOpcional(
        "graficoBateriasCidadesGoias",
        goias,
        "#c5b916"
    );

    criarGraficoOpcional(
        "graficoBateriasCidadesSaoPaulo",
        saoPaulo,
        "#ef233c"
    );
}

function criarGraficoOpcional(
    canvasId,
    dados,
    cor
) {
    if (
        !document.getElementById(
            canvasId
        )
    ) {
        return;
    }

    criarOuAtualizarGrafico(
        canvasId,
        {
            type: "bar",

            data: {
                labels:
                    Object.keys(dados),

                datasets: [{
                    data:
                        Object.values(dados),

                    backgroundColor: cor,
                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor: cor,
                    borderWidth: 1,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 60
                }]
            },

            options:
                opcoesGraficoBarras()
        }
    );
}

function atualizarTop10SitesBahia(contexto) {
    const canvasId =
        "graficoTop10SitesBahia";

    if (
        !document.getElementById(
            canvasId
        )
    ) {
        return;
    }

    const totais =
        Object.create(null);

    contexto.linhasAtual.forEach(
        item => {
            if (
                item.estado !== "BAHIA" ||
                !ehBateria(item) ||
                !item.site
            ) {
                return;
            }

            totais[item.site] =
                (totais[item.site] || 0) +
                item.quantidade;
        }
    );

    const top10 =
        Object.entries(totais)
            .sort(
                (a, b) =>
                    b[1] - a[1]
            )
            .slice(0, 10);

    criarOuAtualizarGrafico(
        canvasId,
        {
            type: "bar",

            data: {
                labels:
                    top10.map(
                        item => item[0]
                    ),

                datasets: [{
                    label:
                        "Quantidade de baterias",

                    data:
                        top10.map(
                            item => item[1]
                        ),

                    backgroundColor:
                        "#f35810",

                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor:
                        "#f35810",

                    borderWidth: 1,

                    hoverBorderColor:
                        "#ffffff",

                    hoverBorderWidth: 2,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 35
                }]
            },

            options:
                opcoesGraficoHorizontal()
        }
    );
}
/* =========================================================
   CABOS — SÃO PAULO
========================================================= */

function atualizarBlocoCabos(contexto) {
    const filtrarCabosGuarulhos =
        linhas =>
            linhas.filter(
                item =>
                    item.estado === "SAO PAULO" &&
                    item.cidade === "GUARULHOS" &&
                    item.tipo.includes("CABO")
            );

    const cabosAtual =
        filtrarCabosGuarulhos(
            contexto.linhasAtual
        );

    const cabosAnterior =
        filtrarCabosGuarulhos(
            contexto.linhasAnterior
        );

    const totalAtual =
        somarQuantidadeProcessada(
            cabosAtual
        );

    const totalAnterior =
        somarQuantidadeProcessada(
            cabosAnterior
        );

    const valorAtual =
        somarValorProcessada(
            cabosAtual
        );

    const valorAnterior =
        somarValorProcessada(
            cabosAnterior
        );

    const ocorrenciasAtual =
        contarIdsProcessados(
            cabosAtual
        );

    const ocorrenciasAnterior =
        contarIdsProcessados(
            cabosAnterior
        );

    alterarTexto(
        "cabosTotal",
        formatarNumero(totalAtual)
    );

    alterarTexto(
        "cabosMediaSemanal",
        formatarNumero(
            calcularMediaSemanalProcessada(
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
        `Período anterior: ${formatarNumero(totalAnterior)}`
    );

    alterarTexto(
        "cabosValorAnterior",
        `Período anterior: ${formatarMoeda(valorAnterior)}`
    );

    alterarTexto(
        "cabosOcorrenciasAnterior",
        `Período anterior: ${formatarNumero(ocorrenciasAnterior)}`
    );

    atualizarGraficoCabosBateriasGuarulhos(
        contexto.linhasAtual
    );

    atualizarTop5BairrosCabosGuarulhos(
        contexto.linhasAtual
    );
}

function atualizarGraficoCabosBateriasGuarulhos(
    linhas
) {
    let totalCabos = 0;
    let totalBaterias = 0;

    linhas.forEach(item => {
        if (
            item.estado !== "SAO PAULO" ||
            item.cidade !== "GUARULHOS"
        ) {
            return;
        }

        if (item.tipo.includes("CABO")) {
            totalCabos += item.quantidade;
        }

        if (item.tipo.includes("BATERIA")) {
            totalBaterias += item.quantidade;
        }
    });

    criarOuAtualizarGrafico(
        "graficoCabosBateriasGuarulhos",
        {
            type: "bar",

            data: {
                labels: [
                    "Cabos",
                    "Baterias"
                ],

                datasets: [{
                    label: "Quantidade furtada",

                    data: [
                        totalCabos,
                        totalBaterias
                    ],

                    backgroundColor: [
                        "#8b5cf6",
                        "#f35810"
                    ],

                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor: [
                        "#a879ff",
                        "#fb7a3c"
                    ],

                    borderWidth: 1,
                    hoverBorderColor: "#ffffff",
                    hoverBorderWidth: 2,

                    borderRadius: 24,
                    borderSkipped: false,
                    maxBarThickness: 100
                }]
            },

            options:
                opcoesGraficoBarras()
        }
    );
}

function atualizarTop5BairrosCabosGuarulhos(
    linhas
) {
    const totaisPorBairro =
        Object.create(null);

    linhas.forEach(item => {
        if (
            item.estado !== "SAO PAULO" ||
            item.cidade !== "GUARULHOS" ||
            !item.tipo.includes("CABO") ||
            !item.bairro
        ) {
            return;
        }

        totaisPorBairro[item.bairro] =
            (totaisPorBairro[item.bairro] || 0) +
            item.quantidade;
    });

    const top5 =
        Object.entries(totaisPorBairro)
            .sort(
                (a, b) =>
                    b[1] - a[1]
            )
            .slice(0, 5);

    criarOuAtualizarGrafico(
        "graficoTop5BairrosCabosGuarulhos",
        {
            type: "bar",

            data: {
                labels:
                    top5.map(
                        item => item[0]
                    ),

                datasets: [{
                    label:
                        "Quantidade de cabos",

                    data:
                        top5.map(
                            item => item[1]
                        ),

                    backgroundColor:
                        "#8b5cf6",

                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor:
                        "#a879ff",

                    borderWidth: 1,
                    hoverBorderColor: "#ffffff",
                    hoverBorderWidth: 2,

                    borderRadius: 20,
                    borderSkipped: false,
                    maxBarThickness: 45
                }]
            },

            options:
                opcoesGraficoHorizontal()
        }
    );
}
    


/* =========================================================
   OCORRÊNCIAS POR ESTADO
========================================================= */

function atualizarOcorrenciasEstados(
    contexto
) {
    const atual =
        contexto.linhasAtual.filter(
            item =>
                estadosMonitorados
                    .includes(item.estado)
        );

    const anterior =
        contexto.linhasAnterior.filter(
            item =>
                estadosMonitorados
                    .includes(item.estado)
        );

    const totalAtual =
        contarIdsProcessados(atual);

    const totalAnterior =
        contarIdsProcessados(anterior);

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
            estado =>
                contarIdsProcessados(
                    atual.filter(
                        item =>
                            item.estado === estado
                    )
                )
        );

    criarOuAtualizarGrafico(
        "graficoOcorrenciasEstados",
        {
            type: "bar",

            data: {
                labels: [
                    "Bahia",
                    "Goiás",
                    "São Paulo"
                ],

                datasets: [{
                    label: "Ocorrências",
                    data: valores,

                    backgroundColor:
                        "#22c55e",

                    hoverBackgroundColor:
                        "#d1d5db",

                    borderColor:
                        "#4ade80",

                    borderWidth: 1,
                    borderRadius: 24,
                    borderSkipped: false,
                    maxBarThickness: 85
                }]
            },

            options:
                opcoesGraficoBarras()
        }
    );
}

/* =========================================================
   RESUMO FINANCEIRO
========================================================= */

function atualizarResumoFinanceiro(
    contexto
) {
    const atual =
        contexto.linhasAtual.filter(
            item =>
                estadosMonitorados
                    .includes(item.estado)
        );

    const anterior =
        contexto.linhasAnterior.filter(
            item =>
                estadosMonitorados
                    .includes(item.estado)
        );

    atualizarFinanceiroEstadoProcessado(
        "BAHIA",
        "valorBahia",
        "valorBahiaVariacao",
        "valorBahiaAnterior",
        atual,
        anterior
    );

    atualizarFinanceiroEstadoProcessado(
        "GOIAS",
        "valorGoias",
        "valorGoiasVariacao",
        "valorGoiasAnterior",
        atual,
        anterior
    );

    atualizarFinanceiroEstadoProcessado(
        "SAO PAULO",
        "valorSaoPaulo",
        "valorSaoPauloVariacao",
        "valorSaoPauloAnterior",
        atual,
        anterior
    );

    const totalAtual =
        somarValorProcessada(atual);

    const totalAnterior =
        somarValorProcessada(anterior);

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

function atualizarFinanceiroEstadoProcessado(
    estado,
    idValor,
    idVariacao,
    idAnterior,
    atual,
    anterior
) {
    const valorAtual =
        somarValorProcessada(
            atual.filter(
                item =>
                    item.estado === estado
            )
        );

    const valorAnterior =
        somarValorProcessada(
            anterior.filter(
                item =>
                    item.estado === estado
            )
        );

    alterarTexto(
        idValor,
        formatarMoeda(valorAtual)
    );

    atualizarVariacao(
        idVariacao,
        valorAtual,
        valorAnterior
    );

    alterarTexto(
        idAnterior,
        `Período anterior: ` +
        `${formatarMoeda(valorAnterior)}`
    );
}

/* =========================================================
   CRIAR OU ATUALIZAR GRÁFICOS
========================================================= */

function criarOuAtualizarGrafico(
    canvasId,
    configuracao
) {
    const canvas =
        document.getElementById(
            canvasId
        );

    if (!canvas) return;

    const existente =
        graficos[canvasId] ||
        Chart.getChart(canvas);

    if (
        existente &&
        existente.config.type ===
            configuracao.type
    ) {
        existente.data.labels =
            configuracao.data.labels;

        existente.data.datasets =
            configuracao.data.datasets;

        existente.options =
            configuracao.options;

        existente.update("none");

        graficos[canvasId] =
            existente;

        return;
    }

    if (existente) {
        existente.destroy();
    }

    graficos[canvasId] =
        new Chart(
            canvas,
            configuracao
        );
}

function opcoesGraficoBarras() {
    return {
        responsive: true,
        maintainAspectRatio: false,

        animation:
            window.innerWidth <= 700
                ? false
                : {
                    duration: 450,
                    easing: "easeOutQuart"
                },

        interaction: {
            mode: "nearest",
            intersect: true
        },

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
                    color: "#9ca3af",
                    precision: 0
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

function opcoesGraficoHorizontal() {
    return {
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
    };
}

/* =========================================================
   CÁLCULOS
========================================================= */

function somarQuantidadeProcessada(
    linhas
) {
    return linhas.reduce(
        (total, item) =>
            total + item.quantidade,
        0
    );
}

function somarValorProcessada(
    linhas
) {
    return linhas.reduce(
        (total, item) =>
            total + item.valor,
        0
    );
}

function contarIdsProcessados(
    linhas
) {
    const ids = new Set();

    linhas.forEach(item => {
        if (item.id) {
            ids.add(item.id);
        }
    });

    return ids.size;
}

function calcularMediaSemanalProcessada(
    linhas,
    intervalo
) {
    const dias =
        Math.max(
            1,

            Math.round(
                (
                    intervalo.fim -
                    intervalo.inicio
                ) /
                86400000
            ) + 1
        );

    return Math.round(
        somarQuantidadeProcessada(
            linhas
        ) /
        Math.max(
            1,
            dias / 7
        )
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

    } else if (percentual < 0) {
        elemento.textContent =
            `▼ ${formatarPercentual(
                Math.abs(percentual)
            )}`;

        elemento.classList.add(
            "variacao-negativa"
        );

    } else {
        elemento.textContent =
            "— 0,0%";
    }
}

/* =========================================================
   DATAS E NÚMEROS
========================================================= */

function converterData(valor) {
    if (!valor) return null;

    if (
        valor instanceof Date &&
        !Number.isNaN(
            valor.getTime()
        )
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

    const brasileira =
        texto.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
        );

    if (brasileira) {
        return new Date(
            Number(brasileira[3]),
            Number(brasileira[2]) - 1,
            Number(brasileira[1])
        );
    }

    const iso =
        new Date(texto);

    return Number.isNaN(
        iso.getTime()
    )
        ? null
        : iso;
}

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
        texto =
            texto
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

    return (
        texto.charAt(0)
            .toUpperCase() +
        texto.slice(1)
    );
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
    const normalizado =
        limparTexto(texto)
            .normalize("NFD")
            .replace(
                /[\u0300-\u036f]/g,
                ""
            )
            .replace(/\s+/g, " ")
            .toLowerCase();

    return (
        aliasesCabecalhos[
            normalizado
        ] ||
        normalizado
    );
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

/* =========================================================
   ANIMAÇÕES AO ROLAR
   NÃO RECALCULA OS GRÁFICOS
========================================================= */

function configurarAnimacoesScroll() {
    const cards =
        document.querySelectorAll(
            ".indicador-card, " +
            ".resumo-card, " +
            ".dashboard-card"
        );

    const graficosPainel =
        document.querySelectorAll(
            ".grafico-painel"
        );

    const blocos =
        document.querySelectorAll(
            ".importacao-sigo, " +
            ".periodo-analise, " +
            ".titulo-bloco"
        );

    cards.forEach(
        (card, indice) => {
            card.classList.add(
                "animar-scroll",

                indice % 2 === 0
                    ? "entrar-esquerda"
                    : "entrar-direita"
            );
        }
    );

    graficosPainel.forEach(
        (grafico, indice) => {
            grafico.classList.add(
                "animar-scroll",

                indice % 2 === 0
                    ? "entrar-esquerda"
                    : "entrar-direita"
            );
        }
    );

    blocos.forEach(bloco => {
        bloco.classList.add(
            "animar-scroll",
            "entrar-baixo"
        );
    });

    const observador =
        new IntersectionObserver(
            entradas => {
                entradas.forEach(
                    entrada => {
                        entrada.target
                            .classList.toggle(
                                "visivel",
                                entrada.isIntersecting
                            );
                    }
                );
            },
            {
                threshold: 0.08,
                rootMargin:
                    "80px 0px 80px 0px"
            }
        );

    document
        .querySelectorAll(
            ".animar-scroll"
        )
        .forEach(elemento => {
            observador.observe(
                elemento
            );
        });
}