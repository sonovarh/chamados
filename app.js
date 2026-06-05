const EMBEDDED_WORKBOOK = {};
let workbookData = {};
let allRows = [];
let filteredRows = [];
let activeQuickFilter = null;
let activeProblemFilter = "__all__";

const SLA_DIAS = 2;
const AUTO_WORKBOOK_FILE = "Chamados RH Sonova (respostas) 05-06-2026.xlsx";

// Padrão Anderson / Igarapé Digital:
// sempre priorizar o último histórico importado manualmente.
// Assim, ao atualizar a tela, fechar e abrir o navegador ou usar o servidor local,
// o dashboard continua com a última base escolhida pelo usuário, e não volta para dados antigos embutidos.
const STORAGE_KEYS = {
  workbook: "sonova_rh_chamados_ultimo_historico_importado_v2",
  meta: "sonova_rh_chamados_ultimo_historico_meta_v2",
  clearMode: "sonova_rh_chamados_base_limpa_v2",
  legacyWorkbook: "sonova_rh_chamados_ultimo_historico_importado_v1",
  legacyMeta: "sonova_rh_chamados_ultimo_historico_meta_v1"
};

const columnAliases = {
  owner: ["Owner", "slav"],
  status: ["Status"],
  observacao: ["Observação", "Observacao"],
  empresa: ["Empresa", "Qual Empresa você está lotado?"],
  data: ["Carimbo de data/hora"],
  email: ["Endereço de e-mail", "Endereco de e-mail"],
  nome: ["NOME"],
  cpf: ["CPF"],
  tipo: ["Tipo de solicitação", "Tipo de solicitacao"],
  descricao: ["Descreva sua solicitação", "Descreva sua solicitacao"],
  anexo: ["Caso precise anexar documentos"],
  urgencia: ["Urgência", "Urgencia"]
};

const CLASSIFIER_GLOSSARY = [
  {
    "id": "acesso_adp",
    "label": "Acesso / Sistema ADP",
    "group": "Sistema",
    "priority": 100,
    "description": "Bloqueio, senha, login, acesso ao ADP/eXpert, portal ou erro de entrada no sistema.",
    "strongTerms": [
      "bloqueei",
      "bloqueado",
      "bloqueada",
      "desbloquear",
      "desbloqueio",
      "senha",
      "login",
      "nao entra",
      "não entra",
      "acesso",
      "acessar",
      "adp",
      "expert",
      "portal"
    ],
    "contextTerms": [
      "folha de pagamento",
      "sistema",
      "plataforma",
      "entrar",
      "erro",
      "usuario",
      "usuário"
    ],
    "negativeTerms": [
      "informe de rendimentos",
      "irrf",
      "malha fina",
      "declaração de ir",
      "declaracao de ir"
    ],
    "examples": [
      "Bloqueei meu acesso do ADP",
      "Alterei a senha e não entra",
      "Preciso desbloquear meu acesso ao portal ADP"
    ]
  },
  {
    "id": "ferias_adp",
    "label": "Férias / ADP",
    "group": "Folha e jornada",
    "priority": 90,
    "description": "Solicitações sobre programação, agendamento, saldo, aviso ou recibo de férias, inclusive erro na plataforma ADP.",
    "strongTerms": [
      "ferias",
      "férias",
      "programar ferias",
      "programar férias",
      "agendamento das minhas ferias",
      "agendamento das minhas férias",
      "recibo de ferias",
      "recibo de férias",
      "saldo de ferias",
      "saldo de férias"
    ],
    "contextTerms": [
      "adp",
      "plataforma",
      "periodo",
      "período",
      "gozo",
      "abono",
      "10 dias",
      "20 dias"
    ],
    "negativeTerms": [],
    "examples": [
      "Não consigo programar minhas férias no ADP",
      "Tirei férias em dois períodos",
      "Preciso consultar saldo de férias"
    ]
  },
  {
    "id": "irrf_informe",
    "label": "IRRF / Informe de rendimentos",
    "group": "Fiscal e declarações",
    "priority": 88,
    "description": "Dúvidas sobre informe de rendimentos, IRRF, declaração de imposto de renda, divergência no pré-preenchido e malha fina.",
    "strongTerms": [
      "irrf",
      "imposto de renda",
      "informe de rendimentos",
      "informe",
      "rendimentos",
      "declaração",
      "declaracao",
      "malha fina",
      "gov",
      "pré-preenchido",
      "pre-preenchido"
    ],
    "contextTerms": [
      "valor",
      "total de rendimentos",
      "empresa",
      "fonte pagadora",
      "receita"
    ],
    "negativeTerms": [
      "bloqueado",
      "bloqueada",
      "senha",
      "login",
      "não entra",
      "nao entra"
    ],
    "examples": [
      "Meu informe está divergente do GOV",
      "Preciso do informe de rendimentos",
      "Dúvida sobre IRRF"
    ]
  },
  {
    "id": "folha_holerite",
    "label": "Folha / Holerite / Pagamento",
    "group": "Folha e jornada",
    "priority": 82,
    "description": "Dúvidas sobre salário, holerite, pagamento, descontos, reembolso, valores pagos ou não pagos.",
    "strongTerms": [
      "folha",
      "holerite",
      "olerite",
      "pagamento",
      "salario",
      "salário",
      "desconto",
      "descontado",
      "cobrança",
      "cobranca",
      "reembolso",
      "valor pago",
      "não recebi",
      "nao recebi"
    ],
    "contextTerms": [
      "competencia",
      "competência",
      "verba",
      "líquido",
      "liquido",
      "adiantamento",
      "13º",
      "decimo terceiro"
    ],
    "negativeTerms": [
      "bloqueado",
      "bloqueada",
      "senha",
      "login"
    ],
    "examples": [
      "Desconto indevido no holerite",
      "Não recebi pagamento",
      "Dúvida sobre folha de pagamento"
    ]
  },
  {
    "id": "ponto_banco_horas",
    "label": "Ponto / Banco de horas",
    "group": "Folha e jornada",
    "priority": 80,
    "description": "Ajustes, dúvidas ou correções de ponto, faltas, atrasos, horas extras e banco de horas.",
    "strongTerms": [
      "ponto",
      "espelho de ponto",
      "banco de horas",
      "hora extra",
      "horas extras",
      "atraso",
      "falta",
      "justificativa",
      "marcação",
      "marcacao",
      "batida"
    ],
    "contextTerms": [
      "jornada",
      "abono",
      "registro",
      "entrada",
      "saida",
      "saída"
    ],
    "negativeTerms": [],
    "examples": [
      "Preciso corrigir meu ponto",
      "Dúvida sobre banco de horas",
      "Falta não abonada"
    ]
  },
  {
    "id": "dependentes_cadastro_familiar",
    "label": "Dependentes / Cadastro familiar",
    "group": "Cadastro",
    "priority": 78,
    "description": "Inclusão, exclusão ou alteração de dependentes e dados familiares para benefícios, IR, saúde ou seguro.",
    "strongTerms": [
      "dependente",
      "dependentes",
      "filha",
      "filho",
      "mãe",
      "mae",
      "pai",
      "cônjuge",
      "conjuge",
      "inclusão de dependente",
      "inclusao de dependente",
      "exclusão de dependente",
      "exclusao de dependente"
    ],
    "contextTerms": [
      "certidão",
      "certidao",
      "documentos",
      "plano",
      "ir",
      "metlife",
      "saúde",
      "saude"
    ],
    "negativeTerms": [],
    "examples": [
      "Quero incluir minha filha",
      "Quem são meus dependentes no MetLife",
      "Excluir dependente do plano"
    ]
  },
  {
    "id": "cadastro_dados",
    "label": "Cadastro / Dados pessoais",
    "group": "Cadastro",
    "priority": 74,
    "description": "Alteração de dados cadastrais, nome, endereço, documentos, banco, e-mail, telefone ou dados pessoais.",
    "strongTerms": [
      "alteração cadastral",
      "alteracao cadastral",
      "cadastro",
      "dados cadastrais",
      "endereço",
      "endereco",
      "telefone",
      "email",
      "e-mail",
      "nome",
      "cpf",
      "rg",
      "banco",
      "conta bancária",
      "conta bancaria"
    ],
    "contextTerms": [
      "corrigir",
      "atualizar",
      "alterar",
      "documento",
      "comprovante"
    ],
    "negativeTerms": [],
    "examples": [
      "Alterar endereço",
      "Corrigir nome no cadastro",
      "Atualizar conta bancária"
    ]
  },
  {
    "id": "plano_saude",
    "label": "Plano de saúde / SulAmérica",
    "group": "Benefícios",
    "priority": 72,
    "description": "Chamados sobre assistência médica, SulAmérica, carteirinha, rede, inclusão/exclusão ou problemas no plano de saúde.",
    "strongTerms": [
      "plano de saude",
      "plano de saúde",
      "saude",
      "saúde",
      "sulamerica",
      "sulamérica",
      "carteirinha",
      "assistência médica",
      "assistencia medica",
      "convênio",
      "convenio"
    ],
    "contextTerms": [
      "dependente",
      "incluir",
      "excluir",
      "rede",
      "coparticipação",
      "coparticipacao"
    ],
    "negativeTerms": [],
    "examples": [
      "Não recebi carteirinha SulAmérica",
      "Alterar dados do plano de saúde",
      "Incluir dependente no plano"
    ]
  },
  {
    "id": "odonto",
    "label": "Odonto",
    "group": "Benefícios",
    "priority": 70,
    "description": "Solicitações de plano odontológico, inclusão, exclusão, carteirinha ou dúvidas do benefício odontológico.",
    "strongTerms": [
      "odonto",
      "odontologico",
      "odontológico",
      "dental"
    ],
    "contextTerms": [
      "plano",
      "carteirinha",
      "dependente",
      "incluir",
      "excluir"
    ],
    "negativeTerms": [],
    "examples": [
      "Solicitação de odonto",
      "Incluir dependente no plano odontológico"
    ]
  },
  {
    "id": "vr_ifood_alelo",
    "label": "VR / iFood / Alelo / Univers",
    "group": "Benefícios",
    "priority": 68,
    "description": "Chamados sobre vale-refeição/alimentação, iFood Benefícios, Alelo, Univers, cartão e saldo de benefícios.",
    "strongTerms": [
      "ifood",
      "i food",
      "alelo",
      "univers",
      "vr",
      "va",
      "vale refeição",
      "vale refeicao",
      "vale alimentação",
      "vale alimentacao",
      "cartão",
      "cartao",
      "beneficio",
      "benefício",
      "saldo"
    ],
    "contextTerms": [
      "crédito",
      "credito",
      "desconto",
      "recarga",
      "aplicativo",
      "acesso ao beneficio"
    ],
    "negativeTerms": [
      "plano de saúde",
      "plano de saude",
      "petlove"
    ],
    "examples": [
      "Não consigo acessar o benefício Univers",
      "Dúvida sobre iFood benefícios",
      "Cartão VR sem saldo"
    ]
  },
  {
    "id": "vale_transporte",
    "label": "Vale-transporte",
    "group": "Benefícios",
    "priority": 66,
    "description": "Solicitações sobre VT, recarga, rota, bilhete, transporte e alteração de vale-transporte.",
    "strongTerms": [
      "vale transporte",
      "vale-transporte",
      "vt",
      "bilhete",
      "transporte",
      "recarga"
    ],
    "contextTerms": [
      "rota",
      "ônibus",
      "onibus",
      "metrô",
      "metro",
      "linha"
    ],
    "negativeTerms": [],
    "examples": [
      "Alterar meu VT",
      "Não caiu recarga do vale-transporte"
    ]
  },
  {
    "id": "metlife_seguro",
    "label": "MetLife / Seguro de vida",
    "group": "Benefícios",
    "priority": 64,
    "description": "Dúvidas sobre seguro de vida, MetLife, beneficiários e dependentes vinculados ao seguro.",
    "strongTerms": [
      "metlife",
      "met life",
      "seguro de vida",
      "seguro",
      "beneficiário",
      "beneficiario"
    ],
    "contextTerms": [
      "dependente",
      "filha",
      "filho",
      "incluir",
      "excluir"
    ],
    "negativeTerms": [
      "petlove"
    ],
    "examples": [
      "Quem são meus dependentes no MetLife",
      "Incluir beneficiário no seguro"
    ]
  },
  {
    "id": "petlove",
    "label": "PetLove",
    "group": "Benefícios",
    "priority": 62,
    "description": "Chamados sobre benefício PetLove, inclusão ou exclusão de pets e plano pet.",
    "strongTerms": [
      "petlove",
      "pet love",
      "plano pet",
      "pets",
      "pet"
    ],
    "contextTerms": [
      "zeus",
      "maria kyara",
      "incluir",
      "excluir",
      "tranquilo"
    ],
    "negativeTerms": [
      "dependente humano"
    ],
    "examples": [
      "Excluir pets do PetLove",
      "Dúvida sobre plano PetLove"
    ]
  },
  {
    "id": "rescisao",
    "label": "Rescisão / Desligamento",
    "group": "Movimentação",
    "priority": 60,
    "description": "Pedidos ou dúvidas relacionados a rescisão, desligamento, verbas rescisórias, TRCT e homologação.",
    "strongTerms": [
      "rescisão",
      "rescisao",
      "desligamento",
      "demissão",
      "demissao",
      "trct",
      "homologação",
      "homologacao",
      "verbas rescisorias",
      "verbas rescisórias"
    ],
    "contextTerms": [
      "fgts",
      "seguro desemprego",
      "aviso prévio",
      "aviso previo"
    ],
    "negativeTerms": [],
    "examples": [
      "Dúvida sobre rescisão",
      "Preciso do TRCT",
      "Verbas rescisórias"
    ]
  },
  {
    "id": "admissao_documentos",
    "label": "Admissão / Documentação",
    "group": "Movimentação",
    "priority": 58,
    "description": "Documentos admissionais, pendências de admissão, exames, cadastro inicial e envio de documentos.",
    "strongTerms": [
      "admissão",
      "admissao",
      "admissional",
      "documentos admissionais",
      "exame admissional",
      "contratação",
      "contratacao"
    ],
    "contextTerms": [
      "enviar documentos",
      "pendência",
      "pendencia",
      "cadastro inicial"
    ],
    "negativeTerms": [],
    "examples": [
      "Pendência de documentos admissionais",
      "Dúvida sobre exame admissional"
    ]
  },
  {
    "id": "declaracoes_comprovantes",
    "label": "Declarações / Comprovantes",
    "group": "Documentos",
    "priority": 54,
    "description": "Solicitações de declarações, comprovantes, carta, documentos e segunda via que não sejam IRRF específico.",
    "strongTerms": [
      "declaração",
      "declaracao",
      "comprovante",
      "carta",
      "segunda via",
      "documento",
      "arquivo",
      "anexo"
    ],
    "contextTerms": [
      "enviar",
      "solicito",
      "preciso",
      "declaração de vínculo",
      "declaracao de vinculo"
    ],
    "negativeTerms": [
      "informe de rendimentos",
      "irrf"
    ],
    "examples": [
      "Preciso de declaração",
      "Solicito comprovante",
      "Enviar segunda via"
    ]
  },
  {
    "id": "outros",
    "label": "Outros não classificados",
    "group": "Outros",
    "priority": 1,
    "description": "Casos que não bateram com segurança em nenhum classificador específico.",
    "strongTerms": [],
    "contextTerms": [],
    "negativeTerms": [],
    "examples": [
      "Mensagem genérica sem termo suficiente"
    ]
  }
];

function normalizeText(value) {
  return String(value ?? "")
    .replace(/\r?\n+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value) {
  return normalizeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function valueByAlias(row, aliasName) {
  const aliases = columnAliases[aliasName] || [];
  const keys = Object.keys(row || {});
  const normalizedKeys = keys.map(k => ({ original: k, normalized: normalizeKey(k) }));

  for (const alias of aliases) {
    const normalizedAlias = normalizeKey(alias);
    const exact = normalizedKeys.find(k => k.normalized === normalizedAlias);
    if (exact) return row[exact.original];
  }

  for (const alias of aliases) {
    const normalizedAlias = normalizeKey(alias);
    const partial = normalizedKeys.find(k => k.normalized.includes(normalizedAlias) || normalizedAlias.includes(k.normalized));
    if (partial) return row[partial.original];
  }

  return "";
}

function normalizeStatus(value) {
  const txt = normalizeText(value);
  return txt || "(vazio)";
}

function normalizeUrgency(value) {
  const n = normalizeKey(value);
  if (!n) return "(vazio)";
  if (n.includes("alta")) return "Alta";
  if (n.includes("media")) return "Média";
  if (n.includes("baixa")) return "Baixa";
  return normalizeText(value) || "(vazio)";
}

function parseExcelDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const txt = normalizeText(value);
  if (!txt) return null;

  const iso = new Date(txt);
  if (!isNaN(iso.getTime())) return iso;

  const br = txt.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (br) {
    const d = Number(br[1]);
    const m = Number(br[2]) - 1;
    let y = Number(br[3]);
    if (y < 100) y += 2000;
    const hh = Number(br[4] || 0);
    const mm = Number(br[5] || 0);
    const ss = Number(br[6] || 0);
    const dt = new Date(y, m, d, hh, mm, ss);
    return isNaN(dt.getTime()) ? null : dt;
  }

  const serial = Number(txt.replace(",", "."));
  if (!isNaN(serial) && serial > 20000 && serial < 80000) {
    const base = new Date(Date.UTC(1899, 11, 30));
    return new Date(base.getTime() + serial * 86400000);
  }

  return null;
}

function dateToIsoDate(value) {
  const d = parseExcelDate(value);
  if (!d) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addDays(date, days) {
  if (!date) return null;
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function daysBetween(start, end) {
  if (!start || !end) return null;
  const one = new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime();
  const two = new Date(end.getFullYear(), end.getMonth(), end.getDate()).getTime();
  return Math.max(0, Math.floor((two - one) / 86400000));
}

function isDone(row) {
  const n = normalizeKey(row.__status);
  return n.includes("concluido") || n.startsWith("0-");
}

function isWaiting(row) {
  const n = normalizeKey(row.__status);
  return n.includes("aguardando") || n.includes("pendente") || n.startsWith("3-");
}

function isHigh(row) {
  return normalizeKey(row.__urgencia) === "alta";
}

function formatDate(value) {
  const d = parseExcelDate(value);
  if (!d) return normalizeText(value);
  return d.toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" });
}

function formatOnlyDate(value) {
  if (!value) return "";
  const d = value instanceof Date ? value : parseExcelDate(value);
  if (!d) return normalizeText(value);
  return d.toLocaleDateString("pt-BR");
}

function classifierTerms(classifier) {
  return [...(classifier.strongTerms || []), ...(classifier.contextTerms || [])];
}

function termHit(text, term) {
  const normalizedTerm = normalizeText(term).toLowerCase();
  if (!normalizedTerm) return false;
  return text.includes(normalizedTerm);
}

function scoreClassifier(row, classifier) {
  if (classifier.id === "outros") {
    return { score: 0, confidence: "Baixa", matchedTerms: [], blocked: false };
  }

  const tipoText = normalizeText(row.__tipo).toLowerCase();
  const descricaoText = normalizeText(`${row.__descricao} ${row.__observacao}`).toLowerCase();
  const baseText = normalizeText(`${row.__tipo} ${row.__descricao} ${row.__observacao}`).toLowerCase();

  const negativeHits = (classifier.negativeTerms || []).filter(term => termHit(descricaoText, term));
  const strongHits = (classifier.strongTerms || []).filter(term => termHit(baseText, term));
  const contextHits = (classifier.contextTerms || []).filter(term => termHit(baseText, term));

  let score = 0;
  score += strongHits.length * 10;
  score += contextHits.length * 3;
  score += Math.min(Number(classifier.priority || 0) / 10, 10);

  // Se a palavra forte está no texto livre do chamado, vale mais do que no tipo escolhido pelo formulário.
  strongHits.forEach(term => {
    if (termHit(descricaoText, term)) score += 6;
    if (termHit(tipoText, term)) score += 2;
  });

  // Bloqueia falsos positivos quando termos negativos aparecem no texto livre.
  const blocked = negativeHits.length > 0 && strongHits.length <= 1;
  if (blocked) score -= 20;

  const confidence = score >= 28 ? "Alta" : score >= 17 ? "Média" : score >= 10 ? "Baixa" : "Muito baixa";
  return {
    score,
    confidence,
    matchedTerms: [...new Set([...strongHits, ...contextHits])],
    blocked
  };
}

function classifyRow(row) {
  const candidates = CLASSIFIER_GLOSSARY
    .filter(c => c.id !== "outros")
    .map(classifier => ({ classifier, result: scoreClassifier(row, classifier) }))
    .filter(item => !item.result.blocked && item.result.score >= 10)
    .sort((a, b) => b.result.score - a.result.score || b.classifier.priority - a.classifier.priority);

  if (!candidates.length) {
    const fallback = row.__tipo && row.__tipo !== "(vazio)" ? row.__tipo : "Outros não classificados";
    return {
      label: fallback,
      group: "Tipo informado",
      confidence: "Baixa",
      score: 0,
      matchedTerms: [],
      classifierId: "fallback_tipo"
    };
  }

  const best = candidates[0];
  return {
    label: best.classifier.label,
    group: best.classifier.group,
    confidence: best.result.confidence,
    score: Math.round(best.result.score),
    matchedTerms: best.result.matchedTerms,
    classifierId: best.classifier.id
  };
}

function getProblemLabel(row) {
  return classifyRow(row).label;
}

function calculateSla(row) {
  const abertura = parseExcelDate(row.__data);
  const prazo = addDays(abertura, SLA_DIAS);
  const hoje = new Date();

  if (!abertura || !prazo) {
    return {
      prazo,
      diasAtePrazo: null,
      atraso: null,
      status: "Sem data"
    };
  }

  if (isDone(row)) {
    return {
      prazo,
      diasAtePrazo: SLA_DIAS,
      atraso: 0,
      status: "Concluído"
    };
  }

  const atraso = hoje > prazo ? daysBetween(prazo, hoje) : 0;
  return {
    prazo,
    diasAtePrazo: SLA_DIAS,
    atraso,
    status: atraso > 0 ? "Fora do SLA" : "Dentro do SLA"
  };
}

function enrichRows(data) {
  const rows = [];
  Object.entries(data).forEach(([sheetName, sheetRows]) => {
    sheetRows.forEach((row, index) => {
      const enriched = {
        ...row,
        __sheet: sheetName,
        __rowNumber: index + 2,
        __owner: normalizeText(valueByAlias(row, "owner")) || "(vazio)",
        __status: normalizeStatus(valueByAlias(row, "status")),
        __empresa: normalizeText(valueByAlias(row, "empresa")) || "(vazio)",
        __data: valueByAlias(row, "data"),
        __dateIso: dateToIsoDate(valueByAlias(row, "data")),
        __email: valueByAlias(row, "email"),
        __nome: valueByAlias(row, "nome"),
        __cpf: valueByAlias(row, "cpf"),
        __tipo: normalizeText(valueByAlias(row, "tipo")) || "(vazio)",
        __descricao: valueByAlias(row, "descricao"),
        __observacao: valueByAlias(row, "observacao"),
        __urgencia: normalizeUrgency(valueByAlias(row, "urgencia"))
      };
      enriched.__classification = classifyRow(enriched);
      enriched.__problem = enriched.__classification.label;
      enriched.__sla = calculateSla(enriched);
      enriched.__search = normalizeText(Object.values(enriched).join(" ")).toLowerCase();
      rows.push(enriched);
    });
  });
  return rows;
}

function countBy(rows, getter) {
  const map = new Map();
  rows.forEach(row => {
    const key = getter(row) || "(vazio)";
    map.set(key, (map.get(key) || 0) + 1);
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
}

function setOptions(selectId, values, current = "__all__") {
  const select = document.getElementById(selectId);
  const defaultLabel = selectId === "sheetFilter" ? "Todas as abas" : selectId === "urgencyFilter" ? "Todas" : "Todos";
  select.innerHTML = `<option value="__all__">${defaultLabel}</option>`;
  values.forEach(v => {
    const option = document.createElement("option");
    option.value = v;
    option.textContent = v;
    select.appendChild(option);
  });
  select.value = values.includes(current) ? current : "__all__";
}

function renderFilters() {
  setOptions("sheetFilter", Object.keys(workbookData), document.getElementById("sheetFilter")?.value || "__all__");
  setOptions("statusFilter", countBy(allRows, r => r.__status).map(x => x[0]), document.getElementById("statusFilter")?.value || "__all__");
  setOptions("ownerFilter", countBy(allRows, r => r.__owner).map(x => x[0]), document.getElementById("ownerFilter")?.value || "__all__");
  setOptions("urgencyFilter", countBy(allRows, r => r.__urgencia).map(x => x[0]), document.getElementById("urgencyFilter")?.value || "__all__");
  setOptions("groupFilter", countBy(allRows, r => r.__classification.group).map(x => x[0]), document.getElementById("groupFilter")?.value || "__all__");
}

function rowMatchesQuickFilter(row) {
  if (!activeQuickFilter) return true;
  if (activeQuickFilter === "done") return isDone(row);
  if (activeQuickFilter === "waiting") return isWaiting(row);
  if (activeQuickFilter === "blankStatus") return row.__status === "(vazio)";
  if (activeQuickFilter === "high") return isHigh(row);
  if (activeQuickFilter === "overdue") return !isDone(row) && row.__sla.atraso > 0;
  if (activeQuickFilter === "insideSla") return !isDone(row) && row.__sla.status === "Dentro do SLA";
  return true;
}

function applyFilters() {
  const sheet = document.getElementById("sheetFilter").value;
  const status = document.getElementById("statusFilter").value;
  const owner = document.getElementById("ownerFilter").value;
  const urgency = document.getElementById("urgencyFilter").value;
  const group = document.getElementById("groupFilter").value;
  const dateFrom = document.getElementById("dateFrom").value;
  const dateTo = document.getElementById("dateTo").value;
  const search = normalizeText(document.getElementById("searchInput").value).toLowerCase();

  filteredRows = allRows.filter(row => {
    if (sheet !== "__all__" && row.__sheet !== sheet) return false;
    if (status !== "__all__" && row.__status !== status) return false;
    if (owner !== "__all__" && row.__owner !== owner) return false;
    if (urgency !== "__all__" && row.__urgencia !== urgency) return false;
    if (group !== "__all__" && row.__classification.group !== group) return false;
    if (activeProblemFilter !== "__all__" && row.__problem !== activeProblemFilter) return false;
    if (!rowMatchesQuickFilter(row)) return false;
    if (dateFrom || dateTo) {
      const d = row.__dateIso || dateToIsoDate(row.__data);
      if (d) {
        if (dateFrom && d < dateFrom) return false;
        if (dateTo && d > dateTo) return false;
      } else {
        return false;
      }
    }
    if (search && !row.__search.includes(search)) return false;
    return true;
  });

  renderDashboard();
}

function kpi(label, value, note, filterKey = "", className = "") {
  const active = filterKey && activeQuickFilter === filterKey ? " active" : "";
  const clickable = filterKey ? " clickable" : "";
  const filterAttr = filterKey ? ` data-filter="${filterKey}" tabindex="0" role="button" aria-label="Filtrar por ${escapeHtml(label)}"` : "";
  return `<article class="kpi ${className}${clickable}${active}"${filterAttr}><span>${label}</span><strong>${value}</strong><small>${note || ""}</small></article>`;
}

function renderKpis(rows) {
  const total = rows.length;
  const done = rows.filter(isDone).length;
  const waiting = rows.filter(isWaiting).length;
  const blankStatus = rows.filter(r => r.__status === "(vazio)").length;
  const high = rows.filter(isHigh).length;
  const overdue = rows.filter(r => !isDone(r) && r.__sla.atraso > 0).length;
  const insideSla = rows.filter(r => !isDone(r) && r.__sla.status === "Dentro do SLA").length;
  const overdueDays = rows.filter(r => !isDone(r) && r.__sla.atraso > 0).map(r => r.__sla.atraso);
  const avgOverdue = overdueDays.length ? Math.round(overdueDays.reduce((a, b) => a + b, 0) / overdueDays.length) : 0;
  const pctDone = total ? Math.round(done / total * 100) : 0;

  document.getElementById("kpiGrid").innerHTML = [
    kpi("Total filtrado", total, "Clique nos cards para refinar", ""),
    kpi("Concluídos", done, `${pctDone}% do total`, "done"),
    kpi("Aguardando", waiting, "Status com pendência", "waiting"),
    kpi("Sem status", blankStatus, "Risco de triagem", "blankStatus"),
    kpi("Alta urgência", high, "Prioridade declarada", "high"),
    kpi("Dentro SLA", insideSla, `Abertura + ${SLA_DIAS} dias`, "insideSla"),
    kpi("Fora SLA", overdue, `Média atraso: ${avgOverdue} dia(s)`, "overdue")
  ].join("");

  document.querySelectorAll(".kpi.clickable").forEach(card => {
    card.addEventListener("click", () => toggleQuickFilter(card.dataset.filter));
    card.addEventListener("keydown", event => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        toggleQuickFilter(card.dataset.filter);
      }
    });
  });
}

function toggleQuickFilter(filterKey) {
  activeQuickFilter = activeQuickFilter === filterKey ? null : filterKey;
  applyFilters();
}

function toggleProblemFilter(label) {
  activeProblemFilter = activeProblemFilter === label ? "__all__" : label;
  applyFilters();
}

function renderBarList(containerId, items, limit = 10, options = {}) {
  const container = document.getElementById(containerId);
  const max = Math.max(...items.map(x => x[1]), 1);
  container.innerHTML = items.slice(0, limit).map(([label, value]) => {
    const pct = Math.round((value / max) * 100);
    const active = options.clickable && activeProblemFilter === label ? " active" : "";
    const clickable = options.clickable ? " clickable" : "";
    const attrs = options.clickable ? ` data-problem="${escapeHtml(label)}" tabindex="0" role="button" aria-label="Filtrar problema ${escapeHtml(label)}"` : "";
    return `<div class="barItem${clickable}${active}"${attrs}>
      <div class="barTop"><b>${escapeHtml(label)}</b><span>${value}</span></div>
      <div class="barTrack"><div class="barFill" style="width:${pct}%"></div></div>
    </div>`;
  }).join("") || "<p>Nenhum registro encontrado.</p>";

  if (options.clickable) {
    container.querySelectorAll(".barItem.clickable").forEach(item => {
      item.addEventListener("click", () => toggleProblemFilter(item.dataset.problem));
      item.addEventListener("keydown", event => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          toggleProblemFilter(item.dataset.problem);
        }
      });
    });
  }
}

function statusBadge(value) {
  const n = normalizeText(value).toLowerCase();
  let cls = "blank";
  if (n.includes("concluido") || n.includes("concluído") || n.startsWith("0-")) cls = "done";
  if (n.includes("aguardando")) cls = "wait";
  return `<span class="badge ${cls}">${escapeHtml(value || "(vazio)")}</span>`;
}

function urgencyBadge(value) {
  const n = normalizeText(value).toLowerCase();
  let cls = "blank";
  if (n === "alta") cls = "alta";
  if (n === "media" || n === "média") cls = "media";
  if (n === "baixa") cls = "baixa";
  return `<span class="badge ${cls}">${escapeHtml(value || "(vazio)")}</span>`;
}

function slaBadge(row) {
  const status = row.__sla.status;
  const cls = status === "Fora do SLA" ? "wait" : status === "Dentro do SLA" ? "done" : "blank";
  const detail = status === "Fora do SLA" ? `${status}: ${row.__sla.atraso} dia(s)` : status;
  return `<span class="badge ${cls}" title="Prazo calculado: abertura + ${SLA_DIAS} dias">${escapeHtml(detail)}</span>`;
}

function renderTable(rows) {
  const table = document.getElementById("dataTable");
  const headers = ["Aba", "Linha", "Abertura", "Prazo SLA", "SLA", "Classificação", "Grupo", "Confiança", "Termos", "Owner", "Status", "Urgência", "Empresa", "Tipo", "Nome", "E-mail", "CPF", "Descrição", "Observação"];
  table.querySelector("thead").innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr>`;
  table.querySelector("tbody").innerHTML = rows.slice(0, 500).map(row => `<tr>
    <td>${escapeHtml(row.__sheet)}</td>
    <td>${row.__rowNumber}</td>
    <td>${escapeHtml(formatDate(row.__data))}</td>
    <td>${escapeHtml(formatOnlyDate(row.__sla.prazo))}</td>
    <td>${slaBadge(row)}</td>
    <td><button class="linkBtn" type="button" data-problem="${escapeHtml(row.__problem)}">${escapeHtml(row.__problem)}</button></td>
    <td>${escapeHtml(row.__classification.group)}</td>
    <td>${escapeHtml(row.__classification.confidence)}</td>
    <td class="terms">${escapeHtml((row.__classification.matchedTerms || []).slice(0, 5).join(", "))}</td>
    <td>${escapeHtml(row.__owner)}</td>
    <td>${statusBadge(row.__status)}</td>
    <td>${urgencyBadge(row.__urgencia)}</td>
    <td>${escapeHtml(row.__empresa)}</td>
    <td>${escapeHtml(row.__tipo)}</td>
    <td>${escapeHtml(row.__nome)}</td>
    <td>${escapeHtml(row.__email)}</td>
    <td>${escapeHtml(row.__cpf)}</td>
    <td class="desc">${escapeHtml(row.__descricao)}</td>
    <td class="desc">${escapeHtml(row.__observacao)}</td>
  </tr>`).join("");

  table.querySelectorAll("button.linkBtn[data-problem]").forEach(button => {
    button.addEventListener("click", () => toggleProblemFilter(button.dataset.problem));
  });

  const filterParts = [];
  if (activeQuickFilter) filterParts.push("card ativo");
  if (activeProblemFilter !== "__all__") filterParts.push(`classificação: ${activeProblemFilter}`);
  const filterText = filterParts.length ? ` Filtro rápido aplicado: ${filterParts.join(" | ")}.` : "";
  document.getElementById("tableSubtitle").textContent = `${rows.length} registros filtrados. Prazo de SLA calculado como abertura + ${SLA_DIAS} dias.${filterText} A tabela mostra até 500 linhas.`;
}

function renderActiveFilterInfo() {
  const info = document.getElementById("activeFilterInfo");
  if (!info) return;
  const parts = [];
  if (activeQuickFilter) parts.push("card de KPI ativo");
  if (activeProblemFilter !== "__all__") parts.push(`classificação: ${activeProblemFilter}`);
  const dateFrom = document.getElementById("dateFrom")?.value;
  const dateTo = document.getElementById("dateTo")?.value;
  if (dateFrom && dateTo) parts.push(`período: ${dateFrom} a ${dateTo}`);
  else if (dateFrom) parts.push(`a partir de: ${dateFrom}`);
  else if (dateTo) parts.push(`até: ${dateTo}`);
  info.textContent = parts.length ? `Filtro ativo: ${parts.join(" | ")}` : "Clique em um card ou em uma classificação para testar se a leitura está boa.";
}


function renderGlossary() {
  const list = document.getElementById("glossaryList");
  if (!list) return;
  const counts = new Map(countBy(allRows, r => r.__problem));
  list.innerHTML = CLASSIFIER_GLOSSARY.filter(c => c.id !== "outros").map(classifier => {
    const count = counts.get(classifier.label) || 0;
    const terms = classifier.strongTerms.slice(0, 10).join(", ");
    return `<details class="glossaryItem">
      <summary><b>${escapeHtml(classifier.label)}</b><span>${count} caso(s)</span></summary>
      <p>${escapeHtml(classifier.description)}</p>
      <p><strong>Grupo:</strong> ${escapeHtml(classifier.group)} | <strong>Prioridade:</strong> ${escapeHtml(classifier.priority)}</p>
      <p><strong>Palavras-chave:</strong> ${escapeHtml(terms)}</p>
      <p><strong>Exemplos:</strong> ${escapeHtml((classifier.examples || []).join(" | "))}</p>
    </details>`;
  }).join("");
}

function downloadClassifiersJson() {
  const payload = {
    versao: "v5_glossario_classificadores",
    regra_sla: `abertura + ${SLA_DIAS} dias`,
    classificadores: CLASSIFIER_GLOSSARY
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "classificadores_chamados_rh_sonova.json";
  a.click();
  URL.revokeObjectURL(url);
}

function renderDashboard() {
  renderKpis(filteredRows);
  renderBarList("problemList", countBy(filteredRows, r => r.__problem), 12, { clickable: true });
  renderBarList("statusList", countBy(filteredRows, r => r.__status), 8);
  renderBarList("ownerList", countBy(filteredRows, r => r.__owner), 8);
  renderBarList("companyList", countBy(filteredRows, r => r.__empresa), 8);
  renderTable(filteredRows);
  renderGlossary();
  renderActiveFilterInfo();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getWorkbookRowCount(data) {
  if (!data || typeof data !== "object") return 0;
  return Object.values(data).reduce((total, rows) => total + (Array.isArray(rows) ? rows.length : 0), 0);
}

function getWorkbookSheetCount(data) {
  if (!data || typeof data !== "object") return 0;
  return Object.keys(data).length;
}

function formatImportedAt(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function saveLastImportedWorkbook(data, fileName) {
  if (!data || typeof data !== "object" || getWorkbookRowCount(data) === 0) return false;

  const meta = {
    fileName: fileName || "Arquivo Excel importado",
    importedAt: new Date().toISOString(),
    totalRows: getWorkbookRowCount(data),
    totalSheets: getWorkbookSheetCount(data),
    fonte: "upload_manual"
  };

  try {
    localStorage.setItem(STORAGE_KEYS.workbook, JSON.stringify(data));
    localStorage.setItem(STORAGE_KEYS.meta, JSON.stringify(meta));
    localStorage.removeItem(STORAGE_KEYS.clearMode);
    localStorage.removeItem(STORAGE_KEYS.legacyWorkbook);
    localStorage.removeItem(STORAGE_KEYS.legacyMeta);
    return true;
  } catch (error) {
    console.warn("Não foi possível salvar o último histórico importado no navegador:", error);
    alert("O arquivo foi importado, mas o navegador não conseguiu salvar o histórico para o próximo acesso. Isso pode ocorrer quando a base está muito grande ou o armazenamento local está bloqueado.");
    return false;
  }
}

function loadLastImportedWorkbookFromStorage() {
  try {
    const rawData = localStorage.getItem(STORAGE_KEYS.workbook);
    if (!rawData) return null;

    const data = JSON.parse(rawData);
    if (!data || typeof data !== "object" || getWorkbookRowCount(data) === 0) return null;

    let meta = {};
    try {
      meta = JSON.parse(localStorage.getItem(STORAGE_KEYS.meta) || "{}");
    } catch (_) {
      meta = {};
    }

    return { data, meta };
  } catch (error) {
    console.warn("Não foi possível carregar o último histórico importado:", error);
    clearLastImportedWorkbook();
    return null;
  }
}

function clearLastImportedWorkbook() {
  try {
    localStorage.removeItem(STORAGE_KEYS.workbook);
    localStorage.removeItem(STORAGE_KEYS.meta);
    localStorage.removeItem(STORAGE_KEYS.legacyWorkbook);
    localStorage.removeItem(STORAGE_KEYS.legacyMeta);
    localStorage.setItem(STORAGE_KEYS.clearMode, "1");
  } catch (error) {
    console.warn("Não foi possível limpar o histórico salvo:", error);
  }
}

function isClearModeActive() {
  try {
    return localStorage.getItem(STORAGE_KEYS.clearMode) === "1";
  } catch (_) {
    return false;
  }
}

function loadWorkbookData(data) {
  workbookData = data || {};
  allRows = enrichRows(workbookData);
  filteredRows = [...allRows];
  renderFilters();
  applyFilters();
}

function workbookToData(workbook) {
  const data = {};
  workbook.SheetNames.forEach(sheetName => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: false });
    data[sheetName] = rows.filter(row => Object.values(row).some(v => String(v).trim() !== ""));
  });
  return data;
}

function readUploadedWorkbook(file) {
  if (typeof XLSX === "undefined") {
    alert("A importação manual pelo navegador depende da biblioteca XLSX online. Para leitura automática real, coloque o Excel na pasta e abra pelo abrir_leitor.bat.");
    return;
  }
  const reader = new FileReader();
  reader.onload = function(event) {
    const bytes = new Uint8Array(event.target.result);
    const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
    const data = workbookToData(workbook);

    loadWorkbookData(data);
    const saved = saveLastImportedWorkbook(data, file.name);
    const totalRows = getWorkbookRowCount(data);
    const totalSheets = getWorkbookSheetCount(data);

    setSourceInfo(
      saved
        ? `Último histórico importado salvo: ${file.name} | ${totalRows} registros | ${totalSheets} aba(s).`
        : `Arquivo importado nesta sessão: ${file.name} | ${totalRows} registros | ${totalSheets} aba(s).`
    );
  };
  reader.readAsArrayBuffer(file);
}

async function tryAutoReadWorkbook() {
  if (isClearModeActive()) {
    loadWorkbookData({});
    setSourceInfo("Base limpa. Nenhuma leitura automática será aplicada até você importar um novo Excel.");
    return;
  }

  // Fluxo preferencial no padrão Anderson / Igarapé Digital:
  // se já houve upload manual, manter esse último histórico como fonte oficial do dashboard.
  const savedWorkbook = loadLastImportedWorkbookFromStorage();
  if (savedWorkbook) {
    loadWorkbookData(savedWorkbook.data);
    const meta = savedWorkbook.meta || {};
    const totalRows = meta.totalRows || getWorkbookRowCount(savedWorkbook.data);
    const totalSheets = meta.totalSheets || getWorkbookSheetCount(savedWorkbook.data);
    const importedAt = formatImportedAt(meta.importedAt);
    const importedAtText = importedAt ? ` | importado em ${importedAt}` : "";
    setSourceInfo(`Último histórico importado: ${meta.fileName || "Arquivo Excel"} | ${totalRows} registros | ${totalSheets} aba(s)${importedAtText}.`);
    return;
  }

  // Fluxo principal: quando aberto pelo abrir_leitor.bat, o server.py lê o XLSX real
  // da pasta e entrega JSON em /api/workbook. Isso evita depender de CDN e evita
  // o bloqueio do navegador contra leitura automática de arquivo local.
  try {
    const response = await fetch("/api/workbook", { cache: "no-store" });
    if (!response.ok) throw new Error("API local não respondeu.");
    const payload = await response.json();
    if (!payload || !payload.sheets || Object.keys(payload.sheets).length === 0) {
      throw new Error("API local não retornou abas válidas.");
    }
    loadWorkbookData(payload.sheets);
    setSourceInfo(`Leitura automática REAL do Excel: ${payload.file || AUTO_WORKBOOK_FILE} | ${payload.totalRows || allRows.length} registros.`);
    return;
  } catch (apiError) {
    console.warn("Falha na API local /api/workbook:", apiError);
  }

  // Fluxo alternativo: se estiver rodando em um servidor estático e a biblioteca XLSX
  // estiver disponível, tenta abrir o arquivo da mesma pasta.
  try {
    if (typeof XLSX === "undefined") throw new Error("Biblioteca XLSX não carregada.");
    const response = await fetch(AUTO_WORKBOOK_FILE, { cache: "no-store" });
    if (!response.ok) throw new Error("Arquivo não encontrado no servidor local.");
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
    loadWorkbookData(workbookToData(workbook));
    setSourceInfo(`Leitura automática do Excel via navegador: ${AUTO_WORKBOOK_FILE}`);
    return;
  } catch (browserError) {
    console.warn("Falha na leitura direta pelo navegador:", browserError);
  }

  // Último recurso: não carregar base antiga. O usuário deve importar a base atual.
  loadWorkbookData({});
  setSourceInfo("Nenhuma base carregada. Importe o Excel atualizado para iniciar o dashboard.");
}

function setSourceInfo(text) {
  const el = document.getElementById("sourceInfo");
  if (el) el.textContent = text;
}

function exportKpiBaseCsv() {
  const headers = [
    "ID_KPI", "Data_Abertura", "Mes_Abertura", "Ano_Abertura", "Aba", "Linha",
    "Owner", "Status", "Status_Grupo", "Urgencia", "Empresa", "Tipo_Formulario",
    "Classificacao", "Grupo_Classificacao", "Confianca", "SLA_Status", "SLA_Prazo", "SLA_Atraso_Dias",
    "Nome", "Email", "CPF", "Descricao", "Observacao"
  ];
  const lines = [headers.join(";")];
  filteredRows.forEach((row, idx) => {
    const d = parseExcelDate(row.__data);
    const mes = d ? String(d.getMonth() + 1).padStart(2, "0") : "";
    const ano = d ? String(d.getFullYear()) : "";
    const statusGrupo = isDone(row) ? "Concluído" : isWaiting(row) ? "Aguardando" : row.__status === "(vazio)" ? "Sem status" : "Em tratamento";
    const values = [
      `${row.__sheet}-${row.__rowNumber}-${idx + 1}`,
      row.__dateIso || dateToIsoDate(row.__data),
      mes,
      ano,
      row.__sheet,
      row.__rowNumber,
      row.__owner,
      row.__status,
      statusGrupo,
      row.__urgencia,
      row.__empresa,
      row.__tipo,
      row.__problem,
      row.__classification.group,
      row.__classification.confidence,
      row.__sla.status,
      formatOnlyDate(row.__sla.prazo),
      row.__sla.atraso ?? 0,
      row.__nome,
      row.__email,
      row.__cpf,
      row.__descricao,
      row.__observacao
    ].map(v => `"${String(v ?? "").replaceAll('"', '""')}"`);
    lines.push(values.join(";"));
  });

  const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "base_kpi_chamados_rh_sonova.csv";
  a.click();
  URL.revokeObjectURL(url);
}

function exportCsv() {
  const headers = ["Aba", "Linha", "Abertura", "Prazo_SLA_Abertura_Mais_2_Dias", "Status_SLA", "Atraso_Dias", "Classificacao", "Grupo_Classificacao", "Confianca", "Termos_Classificacao", "Owner", "Status", "Urgencia", "Empresa", "Tipo", "Nome", "Email", "CPF", "Descricao", "Observacao"];
  const lines = [headers.join(";")];
  filteredRows.forEach(row => {
    const values = [
      row.__sheet, row.__rowNumber, formatDate(row.__data), formatOnlyDate(row.__sla.prazo), row.__sla.status,
      row.__sla.atraso ?? "", row.__problem, row.__classification.group, row.__classification.confidence, (row.__classification.matchedTerms || []).join(", "), row.__owner, row.__status, row.__urgencia, row.__empresa, row.__tipo,
      row.__nome, row.__email, row.__cpf, row.__descricao, row.__observacao
    ].map(v => `"${String(v ?? "").replaceAll('"', '""')}"`);
    lines.push(values.join(";"));
  });
  const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "chamados_rh_filtrado.csv";
  a.click();
  URL.revokeObjectURL(url);
}

document.getElementById("fileInput").addEventListener("change", event => {
  const file = event.target.files?.[0];
  if (file) readUploadedWorkbook(file);
});

document.getElementById("clearDataBtn").addEventListener("click", () => {
  if (!confirm("O último histórico importado será removido da memória deste navegador. Faça o upload do novo Excel para continuar.")) return;
  clearLastImportedWorkbook();
  loadWorkbookData({});
  document.getElementById("fileInput").value = "";
  setSourceInfo("Histórico importado limpo. Faça o upload do novo Excel para salvar uma nova base padrão.");
});

["sheetFilter", "statusFilter", "ownerFilter", "urgencyFilter", "groupFilter", "dateFrom", "dateTo", "searchInput"].forEach(id => {
  document.getElementById(id).addEventListener("input", applyFilters);
  document.getElementById(id).addEventListener("change", applyFilters);
});

document.getElementById("resetBtn").addEventListener("click", () => {
  document.getElementById("sheetFilter").value = "__all__";
  document.getElementById("statusFilter").value = "__all__";
  document.getElementById("ownerFilter").value = "__all__";
  document.getElementById("urgencyFilter").value = "__all__";
  document.getElementById("groupFilter").value = "__all__";
  document.getElementById("dateFrom").value = "";
  document.getElementById("dateTo").value = "";
  document.getElementById("searchInput").value = "";
  activeQuickFilter = null;
  activeProblemFilter = "__all__";
  applyFilters();
});

document.getElementById("exportCsvBtn").addEventListener("click", exportCsv);
const exportKpiBaseBtn = document.getElementById("exportKpiBaseBtn");
if (exportKpiBaseBtn) exportKpiBaseBtn.addEventListener("click", exportKpiBaseCsv);
const exportClassifiersBtn = document.getElementById("exportClassifiersBtn");
if (exportClassifiersBtn) exportClassifiersBtn.addEventListener("click", downloadClassifiersJson);

tryAutoReadWorkbook();
