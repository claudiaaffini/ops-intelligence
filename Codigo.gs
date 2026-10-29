
const NOME_ABA = "Solicitacoes";

const PRAZOS_SLA = {
  "Crítica": 4,
  "Alta": 8,
  "Média": 24,
  "Baixa": 72
};

const STATUS_VALIDOS = [
  "Em aberto",
  "Em análise",
  "Concluída",
  "Cancelada"
];

const COLUNAS = [
  "id",
  "data_abertura",
  "titulo",
  "categoria",
  "equipe",
  "responsavel",
  "prioridade",
  "status",
  "prazo_sla",
  "data_conclusao",
  "descricao"
];

// ========================================
// 1. CONFIGURAÇÃO E ACESSO AOS DADOS
// ========================================

function obterAba() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  if (!planilha) {
    throw new Error("Não foi possível acessar a planilha vinculada.");
  }

  const aba = planilha.getSheetByName(NOME_ABA);

  if (!aba) {
    throw new Error("A aba Solicitacoes não foi encontrada.");
  }

  const cabecalhos = aba
    .getRange(1, 1, 1, COLUNAS.length)
    .getValues()[0]
    .map(valor => String(valor).trim());

  if (cabecalhos.some((valor, i) => valor !== COLUNAS[i])) {
    throw new Error(
      "Os cabeçalhos da planilha não correspondem à estrutura esperada."
    );
  }

  return aba;
}

function testarConexao() {
  const aba = obterAba();

  console.log("Conexão realizada com sucesso!");
  console.log("Planilha: " + aba.getParent().getName());
  console.log("Aba: " + aba.getName());
  console.log("Cabeçalhos: " + COLUNAS.join(", "));
}

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Ops Intelligence");
}

// ========================================
// 2. FUNÇÕES AUXILIARES
// ========================================

function normalizarStatus(status) {
  const valor = String(status || "").trim();

  // Compatibilidade com a versão anterior.
  if (valor === "Pendente") {
    return "Em aberto";
  }

  return valor;
}

function converterData(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  if (typeof valor === "number" && isFinite(valor)) {
    const dataSerial = new Date(
      Date.UTC(1899, 11, 30) + valor * 86400000
    );

    return isNaN(dataSerial.getTime()) ? null : dataSerial;
  }

  if (typeof valor === "string" && valor.trim()) {
    const data = new Date(valor);
    return isNaN(data.getTime()) ? null : data;
  }

  return null;
}

function formatarData(valor) {
  const data = converterData(valor);

  return data
    ? Utilities.formatDate(
        data,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      )
    : "";
}

function obterStatusSla(linha, agora) {
  const status = normalizarStatus(linha[7]);
  const abertura = converterData(linha[1]);
  const prazo = converterData(linha[8]);
  const conclusao = converterData(linha[9]);

  if (!prazo) {
    return "Sem prazo";
  }

  if (status === "Cancelada") {
    return "Cancelada";
  }

  if (status === "Concluída") {
    if (!conclusao) {
      return "Sem data de conclusão";
    }

    return conclusao.getTime() <= prazo.getTime()
      ? "Concluída no prazo"
      : "Concluída com atraso";
  }

  if (!["Em aberto", "Em análise"].includes(status)) {
    return "Status inválido";
  }

  if (agora.getTime() > prazo.getTime()) {
    return "Atrasada";
  }

  if (!abertura) {
    return "Sem data de abertura";
  }

  const duracao = prazo.getTime() - abertura.getTime();
  const restante = prazo.getTime() - agora.getTime();

  if (duracao > 0 && restante <= duracao * 0.25) {
    return "Próxima do vencimento";
  }

  return "Dentro do prazo";
}

function linhaParaObjeto(linha, agora) {
  return {
    id: String(linha[0] || ""),
    dataAbertura: formatarData(linha[1]),
    titulo: String(linha[2] || ""),
    categoria: String(linha[3] || ""),
    equipe: String(linha[4] || ""),
    responsavel: String(linha[5] || ""),
    prioridade: String(linha[6] || ""),
    status: normalizarStatus(linha[7]),
    prazoSla: formatarData(linha[8]),
    dataConclusao: formatarData(linha[9]),
    descricao: String(linha[10] || ""),
    situacaoSla: obterStatusSla(linha, agora)
  };
}

function lerSolicitacoes() {
  const aba = obterAba();
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha <= 1) {
    return [];
  }

  return aba
    .getRange(2, 1, ultimaLinha - 1, COLUNAS.length)
    .getValues()
    .filter(linha => String(linha[0] || "").trim());
}

// ========================================
// 3. CADASTRO DE SOLICITAÇÕES
// ========================================

function cadastrarSolicitacao(dados) {
  if (!dados || typeof dados !== "object") {
    throw new Error("Dados da solicitação inválidos.");
  }

  const camposObrigatorios = [
    "titulo",
    "categoria",
    "equipe",
    "prioridade",
    "descricao"
  ];

  for (const campo of camposObrigatorios) {
    if (
      typeof dados[campo] !== "string" ||
      !dados[campo].trim()
    ) {
      throw new Error("Campo obrigatório inválido: " + campo);
    }
  }

  if (!Object.prototype.hasOwnProperty.call(
    PRAZOS_SLA,
    dados.prioridade
  )) {
    throw new Error("Prioridade inválida.");
  }

  if (
    dados.responsavel != null &&
    typeof dados.responsavel !== "string"
  ) {
    throw new Error("Responsável inválido.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const aba = obterAba();
    const ultimaLinha = aba.getLastRow();

    const ids = ultimaLinha > 1
      ? aba.getRange(2, 1, ultimaLinha - 1, 1)
          .getValues()
          .flat()
      : [];

    const numeros = ids
      .map(id => /^OPS-(\d+)$/.exec(String(id)))
      .filter(resultado => resultado !== null)
      .map(resultado => Number(resultado[1]));

    const proximoNumero = Math.max(0, ...numeros) + 1;
    const id = "OPS-" + String(proximoNumero).padStart(4, "0");

    const dataAbertura = new Date();
    const horasSla = PRAZOS_SLA[dados.prioridade];

    const prazoSla = new Date(
      dataAbertura.getTime() + horasSla * 60 * 60 * 1000
    );

    aba.appendRow([
      id,
      dataAbertura,
      dados.titulo.trim(),
      dados.categoria.trim(),
      dados.equipe.trim(),
      (dados.responsavel || "").trim(),
      dados.prioridade,
      "Em aberto",
      prazoSla,
      "",
      dados.descricao.trim()
    ]);

    SpreadsheetApp.flush();

    console.log("Solicitação cadastrada: " + id);
    return id;

  } finally {
    lock.releaseLock();
  }
}

// ========================================
// 4. LISTAGEM E DETALHAMENTO
// ========================================

function listarSolicitacoes() {
  const agora = new Date();

  return lerSolicitacoes()
    .map(linha => linhaParaObjeto(linha, agora))
    .reverse();
}

function obterSolicitacaoPorId(id) {
  if (typeof id !== "string" || !/^OPS-\d+$/.test(id)) {
    throw new Error("ID inválido.");
  }

  const linha = lerSolicitacoes().find(
    registro => String(registro[0]) === id
  );

  if (!linha) {
    throw new Error("Solicitação não encontrada.");
  }

  return linhaParaObjeto(linha, new Date());
}

// ========================================
// 5. ATUALIZAÇÃO DE STATUS
// ========================================

function atualizarStatusSolicitacao(id, novoStatus) {
  if (typeof id !== "string" || !/^OPS-\d+$/.test(id)) {
    throw new Error("ID inválido.");
  }

  if (!STATUS_VALIDOS.includes(novoStatus)) {
    throw new Error("Status solicitado inválido.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const aba = obterAba();
    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha <= 1) {
      throw new Error("Solicitação não encontrada.");
    }

    const ids = aba
      .getRange(2, 1, ultimaLinha - 1, 1)
      .getValues()
      .flat();

    const indice = ids.findIndex(
      valor => String(valor) === id
    );

    if (indice === -1) {
      throw new Error("Solicitação não encontrada.");
    }

    const numeroLinha = indice + 2;
    const linha = aba
      .getRange(numeroLinha, 1, 1, COLUNAS.length)
      .getValues()[0];

    const statusAnterior = normalizarStatus(linha[7]);

    if (!STATUS_VALIDOS.includes(statusAnterior)) {
      throw new Error(
        "O registro possui um status não reconhecido."
      );
    }

    // Uma mudança de status não altera o prazo original.
    // Ao concluir, registramos a data de conclusão.
    // Ao reabrir ou cancelar, limpamos a data de conclusão.
    let dataConclusao = "";

    if (novoStatus === "Concluída") {
      dataConclusao =
        statusAnterior === "Concluída" &&
        converterData(linha[9])
          ? linha[9]
          : new Date();
    }

    // Atualiza somente as colunas H e J.
    aba.getRange(numeroLinha, 8).setValue(novoStatus);
    aba.getRange(numeroLinha, 10).setValue(dataConclusao);

    SpreadsheetApp.flush();

    const registroAtualizado = aba
      .getRange(numeroLinha, 1, 1, COLUNAS.length)
      .getValues()[0];

    return linhaParaObjeto(registroAtualizado, new Date());

  } finally {
    lock.releaseLock();
  }
}

// ========================================
// 6. INDICADORES DO DASHBOARD
// ========================================

function obterIndicadores() {
  const solicitacoes = lerSolicitacoes();
  const total = solicitacoes.length;

  const concluidas = solicitacoes.filter(
    linha => normalizarStatus(linha[7]) === "Concluída"
  ).length;

  const emAberto = solicitacoes.filter(linha =>
    ["Em aberto", "Em análise"].includes(
      normalizarStatus(linha[7])
    )
  ).length;

  const canceladas = solicitacoes.filter(
    linha => normalizarStatus(linha[7]) === "Cancelada"
  ).length;

  // Cancelamentos não contam como resoluções.
  // A taxa é calculada sobre as solicitações não canceladas.
  const elegiveis = total - canceladas;

  const taxaConclusao = elegiveis > 0
    ? (concluidas / elegiveis) * 100
    : 0;

  return {
    total: total,
    concluidas: concluidas,
    emAberto: emAberto,
    taxaConclusao: Number(taxaConclusao.toFixed(1))
  };
}

// ========================================
// 7. FUNÇÕES DE TESTE
// ========================================

function testarCadastro() {
  const dadosExemplo = {
    titulo: "Solicitar acesso ao sistema interno",
    categoria: "Acesso",
    equipe: "Tecnologia",
    responsavel: "Ana Lima",
    prioridade: "Alta",
    descricao: "Colaborador fictício precisa de acesso."
  };

  const id = cadastrarSolicitacao(dadosExemplo);
  console.log("Teste concluído. ID gerado: " + id);
}

function testarIndicadores() {
  console.log(
    JSON.stringify(obterIndicadores(), null, 2)
  );
}

function testarListagem() {
  console.log(
    JSON.stringify(listarSolicitacoes(), null, 2)
  );
}
