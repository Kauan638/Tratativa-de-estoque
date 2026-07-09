// =====================================
// DADOS GLOBAIS
// =====================================

let workbookHistorico = null;   // workbook XLSX carregado (várias abas = vários itens)
let itensConfig = [];           // [{ aba, sku }] — preenchido pelo usuário
let mapaTratativa = {};         // sku -> { descricao, valorUnitario } (do arquivo opcional)

let resultadoPulmoes = [];      // [{ sku, descricao, endereco, status, primeira, ultima, saldoFinal }]
let resultadoPerdas = [];       // [{ sku, descricao, quantidadePerdida, valorUnitario, perdaEstimada, qtdAjustes }]

let pulmoesFiltrados = [];

// Motivos que entram na conta de "possível perda"
const MOTIVOS_PERDA = [
    "ajuste de endereço",
    "ajuste de armazenagem",
    "inventário"
];

// =====================================
// LOADING
// =====================================

function mostrarLoading(){
    document.getElementById("loadingBox").style.display = "block";
}

function ocultarLoading(){
    document.getElementById("loadingBox").style.display = "none";
}

function atualizarLoading(pct){
    document.getElementById("loadingFill").style.width = pct + "%";
    document.getElementById("loadingPercent").innerText = pct + "%";
}

// =====================================
// TEXTO NORMALIZADO (busca/comparação)
// =====================================

function normalizarTexto(texto){

    return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();

}

function detectarColuna(linhaExemplo, candidatos){

    const chaves = Object.keys(linhaExemplo);

    for(const candidato of candidatos){

        const alvo = normalizarTexto(candidato);

        const encontrada = chaves.find(k => normalizarTexto(k) === alvo);

        if(encontrada) return encontrada;

    }

    for(const candidato of candidatos){

        const alvo = normalizarTexto(candidato);

        const encontrada = chaves.find(k => normalizarTexto(k).includes(alvo));

        if(encontrada) return encontrada;

    }

    return null;

}

// =====================================
// PASSO 1 — CARREGAR ABAS DO HISTÓRICO
// =====================================

function carregarAbas(){

    const arquivo = document.getElementById("arquivoHistorico").files[0];

    document.getElementById("nomeHistorico").innerText =
    arquivo ? arquivo.name : "Nenhum arquivo carregado";

    if(!arquivo) return;

    const leitor = new FileReader();

    leitor.onload = e=>{

        try{

            const dados = new Uint8Array(e.target.result);

            workbookHistorico = XLSX.read(dados, { type: "array" });

            renderizarListaAbas();

        }catch(erro){

            console.error(erro);

            alert("Erro ao ler o arquivo: " + erro.message);

        }

    };

    leitor.readAsArrayBuffer(arquivo);

}

function renderizarListaAbas(){

    const container = document.getElementById("listaAbas");

    container.innerHTML = "";

    workbookHistorico.SheetNames.forEach(nomeAba=>{

        const linha = document.createElement("div");

        linha.className = "aba-linha";

        linha.innerHTML = `
            <div class="aba-nome">📑 ${nomeAba}</div>
            <input type="text" placeholder="SKU / Código do item" id="sku-${cssEscape(nomeAba)}">
            <input type="text" placeholder="Descrição (opcional — vem do arquivo de Tratativa se você subir)" id="desc-${cssEscape(nomeAba)}">
        `;

        container.appendChild(linha);

    });

}

function cssEscape(str){

    return str.replace(/[^a-zA-Z0-9_-]/g, "_");

}

// =====================================
// PASSO 2 — LER TRATATIVA DE ESTOQUE (opcional)
// =====================================

function lerPlanilhaGenerica(arquivo){

    return new Promise((resolve, reject)=>{

        const leitor = new FileReader();

        leitor.onload = e=>{

            try{

                const dados = new Uint8Array(e.target.result);

                const workbook = XLSX.read(dados, { type: "array" });

                const aba = workbook.SheetNames[0];

                const linhas = XLSX.utils.sheet_to_json(
                    workbook.Sheets[aba],
                    { defval: "" }
                );

                resolve(linhas);

            }catch(erro){

                reject(erro);

            }

        };

        leitor.onerror = () => reject(new Error("Falha ao ler " + arquivo.name));

        leitor.readAsArrayBuffer(arquivo);

    });

}

async function lerTratativa(){

    const arquivo = document.getElementById("arquivoTratativa").files[0];

    document.getElementById("nomeTratativa").innerText =
    arquivo ? arquivo.name : "Nenhum arquivo carregado";

    mapaTratativa = {};

    if(!arquivo) return;

    const linhas = await lerPlanilhaGenerica(arquivo);

    if(!linhas.length) return;

    const colSku = detectarColuna(linhas[0], [
        "sku", "codigo", "código", "codigo do produto"
    ]);

    const colDescricao = detectarColuna(linhas[0], [
        "descricao", "descrição", "produto"
    ]);

    const colValor = detectarColuna(linhas[0], [
        "valor unitario", "valor unitário", "valor", "custo"
    ]);

    if(!colSku){

        console.warn("Tratativa: não achei coluna de SKU. Colunas:", Object.keys(linhas[0]));

        return;

    }

    linhas.forEach(linha=>{

        const sku = String(linha[colSku] || "").trim();

        if(!sku) return;

        mapaTratativa[sku] = {

            descricao: colDescricao ? String(linha[colDescricao] || "").trim() : "",
            valorUnitario: colValor ? Number(linha[colValor]) || 0 : 0

        };

    });

}

document
.getElementById("arquivoTratativa")
.addEventListener("change", lerTratativa);

// =====================================
// PARSER DO HISTÓRICO (por aba)
// =====================================

// Endereço vem como "End. <CODIGO> <TIPO>" onde TIPO é
// APANHA, PULMAO ou AVARIA — código pode ter letras, números,
// pontos e caracteres especiais (ex: BOX, J&T, AVA.1.0.4175).
const REGEX_ENDERECO = /^End\.\s*(\S+)\s+(APANHA|PULMAO|AVARIA)/i;

function parseLinhaHistorico(row){

    // row = [data, tipo(S/E), endereco, motivo, saldoAnterior, quantidade, saldoResultante, embalagem, observacao]

    const [data, tipo, enderecoTexto, motivo, saldoAnterior, quantidade, saldoResultante] = row;

    if(!enderecoTexto) return null;

    const match = String(enderecoTexto).trim().match(REGEX_ENDERECO);

    if(!match) return null;

    return {

        data,
        tipo,
        codigoEndereco: match[1],
        tipoEndereco: match[2].toUpperCase(),
        motivo: motivo || "",
        saldoAnterior: Number(saldoAnterior) || 0,
        quantidade: Number(quantidade) || 0,
        saldoResultante: Number(saldoResultante) || 0

    };

}

function processarAba(nomeAba, sku, descricaoManual){

    const sheet = workbookHistorico.Sheets[nomeAba];

    const linhasBrutas = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: null,
        raw: true
    });

    const registros = linhasBrutas
    .map(parseLinhaHistorico)
    .filter(Boolean);

    // ---- pulmões: agrupa por endereço, guarda 1ª/última movimentação e saldo final ----
    const pulmoesPorEndereco = {};

    registros
    .filter(r => r.tipoEndereco === "PULMAO")
    .forEach(r=>{

        const chave = r.codigoEndereco;

        if(!pulmoesPorEndereco[chave]){

            pulmoesPorEndereco[chave] = {

                endereco: chave,
                primeira: r.data,
                ultima: r.data,
                saldoFinal: r.saldoResultante

            };

        }else{

            const p = pulmoesPorEndereco[chave];

            if(r.data && (!p.primeira || r.data < p.primeira)) p.primeira = r.data;
            if(r.data && (!p.ultima || r.data >= p.ultima)){

                p.ultima = r.data;
                p.saldoFinal = r.saldoResultante;

            }

        }

    });

    const infoTratativa = mapaTratativa[sku];

    const descricaoFinal =
    (infoTratativa && infoTratativa.descricao) ||
    descricaoManual ||
    "";

    const pulmoesDoItem = Object.values(pulmoesPorEndereco).map(p => ({

        sku,
        descricao: descricaoFinal,
        endereco: p.endereco,
        status: p.saldoFinal > 0 ? "ativo" : "antigo",
        primeira: p.primeira,
        ultima: p.ultima,
        saldoFinal: p.saldoFinal

    }));

    // ---- possíveis perdas: soma dos deltas negativos em motivos de ajuste/inventário ----
    let quantidadePerdida = 0;
    let qtdAjustes = 0;

    registros.forEach(r=>{

        const motivoNormalizado = normalizarTexto(r.motivo);

        const ehMotivoDePerda = MOTIVOS_PERDA.some(
            m => motivoNormalizado.includes(m)
        );

        if(!ehMotivoDePerda) return;

        const delta = r.saldoResultante - r.saldoAnterior;

        if(delta < 0){

            quantidadePerdida += Math.abs(delta);
            qtdAjustes++;

        }

    });

    const valorUnitario = infoTratativa ? infoTratativa.valorUnitario : 0;

    const perda = {

        sku,
        descricao: descricaoFinal,
        quantidadePerdida,
        valorUnitario,
        perdaEstimada: valorUnitario ? quantidadePerdida * valorUnitario : null,
        qtdAjustes

    };

    return { pulmoes: pulmoesDoItem, perda };

}

// =====================================
// PROCESSAMENTO PRINCIPAL
// =====================================

async function processarTudo(){

    if(!workbookHistorico){

        alert("Suba o arquivo de histórico primeiro.");

        return;

    }

    // valida que todo mundo preencheu o SKU
    const abas = workbookHistorico.SheetNames;

    itensConfig = [];

    for(const aba of abas){

        const sku = document.getElementById(`sku-${cssEscape(aba)}`).value.trim();

        if(!sku){

            alert(`Preencha o SKU da aba "${aba}" antes de processar.`);

            return;

        }

        const descricaoManual = document.getElementById(`desc-${cssEscape(aba)}`).value.trim();

        itensConfig.push({ aba, sku, descricaoManual });

    }

    mostrarLoading();
    atualizarLoading(10);

    // garante que a tratativa (se selecionada mas ainda não lida) está carregada
    if(document.getElementById("arquivoTratativa").files[0] && Object.keys(mapaTratativa).length === 0){

        await lerTratativa();

    }

    atualizarLoading(30);

    resultadoPulmoes = [];
    resultadoPerdas = [];

    itensConfig.forEach((item, idx)=>{

        const { pulmoes, perda } = processarAba(item.aba, item.sku, item.descricaoManual);

        resultadoPulmoes.push(...pulmoes);
        resultadoPerdas.push(perda);

        atualizarLoading(30 + Math.round(60 * (idx + 1) / itensConfig.length));

    });

    pulmoesFiltrados = [...resultadoPulmoes];

    atualizarLoading(100);

    atualizarKPIs();
    renderizarPulmoes();
    renderizarPerdas();

    setTimeout(ocultarLoading, 300);

}

// =====================================
// KPIs
// =====================================

function formatarData(data){

    if(!data) return "-";

    const d = data instanceof Date ? data : new Date(data);

    if(isNaN(d)) return "-";

    return d.toLocaleDateString("pt-BR");

}

function atualizarKPIs(){

    const totalItens = itensConfig.length;

    const enderecosUnicos = new Set(resultadoPulmoes.map(p => p.sku + "|" + p.endereco));

    const ativos = resultadoPulmoes.filter(p => p.status === "ativo").length;
    const antigos = resultadoPulmoes.filter(p => p.status === "antigo").length;

    const perdaQtdTotal = resultadoPerdas.reduce((s,p)=> s + p.quantidadePerdida, 0);

    const temAlgumValor = resultadoPerdas.some(p => p.perdaEstimada !== null);

    const perdaValorTotal = resultadoPerdas.reduce(
        (s,p)=> s + (p.perdaEstimada || 0), 0
    );

    document.getElementById("kpiItens").textContent = totalItens;
    document.getElementById("kpiPulmoes").textContent = enderecosUnicos.size;
    document.getElementById("kpiAtivos").textContent = ativos;
    document.getElementById("kpiAntigos").textContent = antigos;
    document.getElementById("kpiPerdaQtd").textContent = perdaQtdTotal.toLocaleString("pt-BR");

    document.getElementById("kpiPerdaValor").textContent =
    temAlgumValor
    ? perdaValorTotal.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })
    : "— (envie a Tratativa)";

}

// =====================================
// FILTROS
// =====================================

function aplicarFiltros(){

    const busca = normalizarTexto(document.getElementById("fSku").value);
    const status = document.getElementById("fStatusPulmao").value;

    pulmoesFiltrados = resultadoPulmoes.filter(p=>{

        const bateBusca =
        !busca ||
        normalizarTexto(p.sku).includes(busca) ||
        normalizarTexto(p.descricao).includes(busca) ||
        normalizarTexto(p.endereco).includes(busca);

        const bateStatus =
        !status || p.status === status;

        return bateBusca && bateStatus;

    });

    renderizarPulmoes();

}

function limparFiltros(){

    document.getElementById("fSku").value = "";
    document.getElementById("fStatusPulmao").value = "";

    pulmoesFiltrados = [...resultadoPulmoes];

    renderizarPulmoes();

}

function filtrarStatus(status){

    document.getElementById("fStatusPulmao").value = status;

    aplicarFiltros();

}

// =====================================
// RENDERIZAÇÃO — PULMÕES
// =====================================

function renderizarPulmoes(){

    const tbody = document.getElementById("tbodyPulmoes");

    tbody.innerHTML = "";

    if(!pulmoesFiltrados.length){

        tbody.innerHTML = `
        <tr>
            <td colspan="7" style="text-align:center; padding:24px; color:var(--text-muted);">
                Nenhum pulmão encontrado. Processe o histórico primeiro.
            </td>
        </tr>
        `;

        return;

    }

    pulmoesFiltrados.forEach(p=>{

        const badge =
        p.status === "ativo"
        ? `<span class="badge-status badge-ativo">🟢 Ativo</span>`
        : `<span class="badge-status badge-antigo">⚪ Antigo</span>`;

        tbody.innerHTML += `
        <tr>
            <td>${p.sku}</td>
            <td>${p.descricao || "<span class='sem-valor'>—</span>"}</td>
            <td>${p.endereco}</td>
            <td>${badge}</td>
            <td>${formatarData(p.primeira)}</td>
            <td>${formatarData(p.ultima)}</td>
            <td>${p.saldoFinal}</td>
        </tr>
        `;

    });

}

// =====================================
// RENDERIZAÇÃO — PERDAS
// =====================================

function renderizarPerdas(){

    const tbody = document.getElementById("tbodyPerdas");

    tbody.innerHTML = "";

    if(!resultadoPerdas.length){

        tbody.innerHTML = `
        <tr>
            <td colspan="6" style="text-align:center; padding:24px; color:var(--text-muted);">
                Nenhum dado. Processe o histórico primeiro.
            </td>
        </tr>
        `;

        return;

    }

    resultadoPerdas.forEach(p=>{

        const valorUnitarioTexto =
        p.valorUnitario
        ? p.valorUnitario.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })
        : `<span class="sem-valor">sem valor</span>`;

        const perdaTexto =
        p.perdaEstimada !== null
        ? p.perdaEstimada.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })
        : `<span class="sem-valor">envie a Tratativa</span>`;

        tbody.innerHTML += `
        <tr>
            <td>${p.sku}</td>
            <td>${p.descricao || "<span class='sem-valor'>—</span>"}</td>
            <td>${p.quantidadePerdida.toLocaleString("pt-BR")}</td>
            <td>${valorUnitarioTexto}</td>
            <td>${perdaTexto}</td>
            <td>${p.qtdAjustes}</td>
        </tr>
        `;

    });

}

// =====================================
// EXPORTAR EXCEL
// =====================================

function exportarExcel(){

    if(!resultadoPulmoes.length){

        alert("Processe o histórico primeiro.");

        return;

    }

    const wb = XLSX.utils.book_new();

    const abaPulmoes = resultadoPulmoes.map(p => ({

        SKU: p.sku,
        Descricao: p.descricao,
        EnderecoPulmao: p.endereco,
        Status: p.status === "ativo" ? "Ativo" : "Antigo/Encerrado",
        PrimeiraMovimentacao: formatarData(p.primeira),
        UltimaMovimentacao: formatarData(p.ultima),
        SaldoFinal: p.saldoFinal

    }));

    const abaPerdas = resultadoPerdas.map(p => ({

        SKU: p.sku,
        Descricao: p.descricao,
        QuantidadePerdida: p.quantidadePerdida,
        ValorUnitario: p.valorUnitario || "",
        PerdaEstimadaRS: p.perdaEstimada !== null ? p.perdaEstimada.toFixed(2) : "sem valor",
        QtdAjustes: p.qtdAjustes

    }));

    XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.json_to_sheet(abaPulmoes),
        "Pulmoes"
    );

    XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.json_to_sheet(abaPerdas),
        "Perdas"
    );

    XLSX.writeFile(wb, "Tratativa_Estoque_Consolidado.xlsx");

}
