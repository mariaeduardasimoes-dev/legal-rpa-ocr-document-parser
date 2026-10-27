/**
 * Unified Legal RPA & OCR Document Parser for Hearings and Technical Inspections
 * Version v9.4 Maximum Precision - Heuristic ERP/SAP Data Extraction & Sanitization
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗓️ Funções da Planilha')
    .addItem('1. Agendar Compromissos no Calendar', 'agendarCompromissos')
    .addSeparator()
    .addItem('2. Preencher Dados de Perícia com IA/OCR', 'processarLinhasComIA')
    .addToUi();
}

// ================================================================== //
// SECTION 1: EXPERT INSPECTION PROCESSING (MULTI-PDF + OCR + ERP ENGINE)
// ================================================================== //

function eTipoPericia(tipo) {
  if (!tipo) return false;
  const t = tipo.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "");
  return ["insalubridade", "periculosidade", "periciamedicainterna", "periciamedicaexterna", "pericia"].some(p => t.includes(p));
}

function processarLinhasComIA() {
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (["aoEditarPlanilhaAuto", "varreduraPeriodicaAuto"].includes(t.getHandlerFunction())) {
        ScriptApp.deleteTrigger(t);
      }
    });
  } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Audiencias");
  if (!aba) {
    ss.toast("Aba 'Audiencias' não encontrada!", "❌ Erro", 8);
    return;
  }

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) {
    ss.toast("Não há dados na planilha para processar.", "⚠️ Aviso", 6);
    return;
  }

  ss.toast("Verificando linhas com campos pendentes ou para higienizar...", "🔍 Iniciando", 5);
  const numColunas = Math.max(aba.getLastColumn(), 17);
  const dados = aba.getRange(2, 1, ultimaLinha - 1, numColunas).getValues();
  let linhasParaProcessar = [];

  dados.forEach((linha, index) => {
    const tipoCompromisso = linha[5]; // Coluna F
    const cargo = linha[9];           // Coluna J
    const cc = linha[10];             // Coluna K
    const horario = linha[11];        // Coluna L
    const admissao = linha[12];       // Coluna M

    const cargoStr = cargo ? cargo.toString().trim() : "";
    const ccStr = cc ? cc.toString().trim() : "";

    const cargoPrecisaAjuste = !cargoStr || cargoStr === "N/A" || /^associad[oa]$/i.test(cargoStr) || /descri[cç][aã]o\s+do\s+turno/i.test(cargoStr);
    const ccPrecisaAjuste = !ccStr || ccStr === "N/A" || /^manuten[cç][aã]o\s+de$/i.test(ccStr) || /^1000[A-Z]\d{6}/i.test(ccStr);

    const temCampoPendente = cargoPrecisaAjuste || ccPrecisaAjuste || [horario, admissao].some(
      val => !val || val.toString().trim() === "" || val.toString().trim() === "N/A"
    );

    if (eTipoPericia(tipoCompromisso) && temCampoPendente) {
      const numLinha = index + 2;
      linhasParaProcessar.push(numLinha);
      aba.getRange(numLinha, 17).setValue("⏳ Na fila...");
    }
  });

  SpreadsheetApp.flush();

  if (linhasParaProcessar.length === 0) {
    ss.toast("Todas as perícias já estão preenchidas!", "ℹ️ Nada pendente", 8);
    return;
  }

  ss.toast(`Processando ${linhasParaProcessar.length} perícia(s)... Acompanhe na Coluna Q!`, "🚀 Executando", 10);
  linhasParaProcessar.forEach(numLinha => {
    executarLogicaIA(aba, numLinha);
  });
  ss.toast("Processamento concluído! Verifique os dados e a Coluna Q.", "✅ Finalizado!", 10);
}

function ejecutarLogicaIA(aba, numLinha) {
  const statusCell = aba.getRange(numLinha, 17); // Coluna Q
  const nomeReclamante = aba.getRange(numLinha, 1).getValue(); // Coluna A
  const numProcesso = aba.getRange(numLinha, 2).getValue();    // Coluna B

  if (!nomeReclamante) return statusCell.setValue("❌ Reclamante Vazio");
  statusCell.setValue("⏳ Buscando pasta...");
  SpreadsheetApp.flush();

  try {
    const props = PropertiesService.getScriptProperties();
    const PASTA_ORIGINAL_ID = props.getProperty('PRIMARY_FOLDER_ID') || "YOUR_PRIMARY_DRIVE_FOLDER_ID";
    const PASTA_BACKUP_ID = props.getProperty('BACKUP_FOLDER_ID') || "YOUR_BACKUP_DRIVE_FOLDER_ID";

    let itemEncontrado = encontrarItemPorReclamante(PASTA_ORIGINAL_ID, nomeReclamante, numProcesso);
    if (!itemEncontrado) {
      itemEncontrado = encontrarItemPorReclamante(PASTA_BACKUP_ID, nomeReclamante, numProcesso);
    }

    if (!itemEncontrado) return statusCell.setValue("❌ Pasta/PDF não encontrado");

    statusCell.setValue("⏳ Lendo documentos...");
    SpreadsheetApp.flush();

    const dadosExtraidos = extrairDadosMultiPdf(itemEncontrado, statusCell);
    if (!dadosExtraidos.leuAlgumPdf) {
      return statusCell.setValue("❌ PDF ilegível ou vazio");
    }

    const atuaisLimpos = limparDadosExtraidos({
      cargo: aba.getRange(numLinha, 10).getValue() ? aba.getRange(numLinha, 10).getValue().toString() : "N/A",
      centro_de_custo: aba.getRange(numLinha, 11).getValue() ? aba.getRange(numLinha, 11).getValue().toString() : "N/A",
      horario_de_trabalho: aba.getRange(numLinha, 12).getValue() ? aba.getRange(numLinha, 12).getValue().toString() : "N/A",
      data_admissao: aba.getRange(numLinha, 13).getValue() ? formatarData(aba.getRange(numLinha, 13).getValue()) : "N/A"
    });

    const ehValido = (v) => v && v.toString().trim() !== "" && v.toString().trim() !== "N/A";
    const ehCargoValido = (v) => ehValido(v) && !/^associad[oa]$/i.test(v.toString().trim());

    const finalCargo = ehValido(dadosExtraidos.cargo) ? dadosExtraidos.cargo : (ehCargoValido(atuaisLimpos.cargo) ? atuaisLimpos.cargo : "N/A");
    const finalCC = ehValido(dadosExtraidos.centro_de_custo) ? dadosExtraidos.centro_de_custo : (ehValido(atuaisLimpos.centro_de_custo) ? atuaisLimpos.centro_de_custo : "N/A");
    const finalHor = ehValido(dadosExtraidos.horario_de_trabalho) ? dadosExtraidos.horario_de_trabalho : (ehValido(atuaisLimpos.horario_de_trabalho) ? atuaisLimpos.horario_de_trabalho : "N/A");
    const finalAdm = ehValido(dadosExtraidos.data_admissao) ? dadosExtraidos.data_admissao : (ehValido(atuaisLimpos.data_admissao) ? atuaisLimpos.data_admissao : "N/A");

    aba.getRange(numLinha, 10).setValue(finalCargo);
    aba.getRange(numLinha, 11).setValue(finalCC);
    aba.getRange(numLinha, 12).setValue(finalHor);
    const cellAdmissao = aba.getRange(numLinha, 13);
    cellAdmissao.setNumberFormat("@");
    cellAdmissao.setValue(finalAdm);

    const totalEncontrados = [finalCargo, finalCC, finalHor, finalAdm].filter(v => v !== "N/A").length;
    statusCell.setValue(totalEncontrados === 4 ? "✅ Processado (4/4)" : `✅ Processado (${totalEncontrados}/4)`);
    SpreadsheetApp.flush();

  } catch (err) {
    statusCell.setValue(`❌ Erro: ${err.message.substring(0, 80)}`);
    Logger.log(`Erro na linha ${numLinha}: ${err.stack}`);
  }
}

/**
 * Fuzzy Search with Scoring Algorithm across Google Drive Folders
 */
function encontrarItemPorReclamante(pastaId, nomeCompleto, numProcesso) {
  try {
    const pastaRaiz = DriveApp.getFolderById(pastaId);
    const procDigitos = numProcesso ? numProcesso.toString().replace(/[^0-9]/g, "") : "";
    const procInicio = procDigitos.length >= 7 ? procDigitos.substring(0, 7) : "";

    const palavrasOriginais = nomeCompleto.toString().trim().split(/\s+/)
      .filter(p => p.length >= 3 && !["dos", "das", "del", "com", "de", "da"].includes(p.toLowerCase()));

    const palavras = nomeCompleto.toString().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9 ]/g, " ")
      .split(/\s+/)
      .filter(p => p.length >= 3 && !["dos", "das", "del", "com", "de", "da"].includes(p));

    if (palavras.length === 0 && !procInicio) return null;

    if (palavrasOriginais.length > 0) {
      const primeiroNomeEscapado = palavrasOriginais[0].replace(/'/g, "\\'");
      try {
        const buscaRapida = pastaRaiz.searchFolders(`title contains '${primeiroNomeEscapado}' and trashed = false`);
        let melhorRapido = null;
        let maiorPontosRapido = 0;

        while (buscaRapida.hasNext()) {
          const p = buscaRapida.next();
          const nomeNorm = p.getName().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
          const nomeDigitos = p.getName().replace(/[^0-9]/g, "");
          let pts = 30;
          if (procInicio && nomeDigitos.includes(procInicio)) pts += 150;
          palavras.forEach(w => { if (nomeNorm.includes(w)) pts += 25; });
          if (pts > maiorPontosRapido) {
            maiorPontosRapido = pts;
            melhorRapido = p;
          }
        }
        if (maiorPontosRapido >= 75 && melhorRapido) return melhorRapido;
      } catch (eSearch) {}
    }

    let melhorItem = null;
    let maiorPontuacao = 0;
    let contador = 0;
    const LIMITE_PASTAS = 500;
    const subpastas = pastaRaiz.getFolders();

    while (subpastas.hasNext() && contador < LIMITE_PASTAS) {
      contador++;
      const pasta = subpastas.next();
      const nomePastaNorm = pasta.getName().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      const nomePastaDigitos = pasta.getName().replace(/[^0-9]/g, "");
      let pontos = 0;

      if (procInicio && nomePastaDigitos.includes(procInicio)) pontos += 150;
      if (palavras.length > 0) {
        const primeiroNome = palavras[0];
        const ultimoNome = palavras[palavras.length - 1];
        if (nomePastaNorm.includes(primeiroNome) && nomePastaNorm.includes(ultimoNome)) {
          pontos += 100;
        } else if (palavras.length > 1 && nomePastaNorm.includes(primeiroNome) && nomePastaNorm.includes(palavras[1])) {
          pontos += 90;
        } else if (nomePastaNorm.includes(primeiroNome)) {
          pontos += 40;
        }
        palavras.forEach(p => { if (nomePastaNorm.includes(p)) pontos += 10; });
      }
      if (pontos > maiorPontuacao) {
        maiorPontuacao = pontos;
        melhorItem = pasta;
      }
    }

    if (maiorPontuacao >= 50) return melhorItem;

    if (palavrasOriginais.length > 0) {
      const primeiroNomeEscapado = palavrasOriginais[0].replace(/'/g, "\\'");
      const arquivos = pastaRaiz.searchFiles(`title contains '${primeiroNomeEscapado}' and trashed = false`);
      if (arquivos.hasNext()) return arquivos.next();
    }
  } catch (e) {
    Logger.log("Erro ao buscar item: " + e.toString());
  }
  return null;
}

/**
 * Scans up to 10 prioritized PDFs inside the claimant folder
 */
function extrairDadosMultiPdf(item, statusCell) {
  let listaPdfIds = [];
  if (typeof item.getFiles === "function") {
    listaPdfIds = listarPdfsDaPastaOrdenados(item);
  } else if (typeof item.getMimeType === "function") {
    const mime = item.getMimeType();
    if (mime === "application/vnd.google-apps.shortcut") {
      try {
        const targetId = item.getTargetId();
        try {
          const pastaAlvo = DriveApp.getFolderById(targetId);
          listaPdfIds = listarPdfsDaPastaOrdenados(pastaAlvo);
        } catch (e) {
          listaPdfIds = [targetId];
        }
      } catch (err) {}
    } else if (mime === MimeType.PDF || item.getName().toLowerCase().endsWith(".pdf")) {
      listaPdfIds = [item.getId()];
    }
  }

  const acumulado = {
    cargo: "N/A",
    centro_de_custo: "N/A",
    horario_de_trabalho: "N/A",
    data_admissao: "N/A",
    leuAlgumPdf: false
  };

  let achouAssociadoFallback = false;
  const maxPdfs = Math.min(listaPdfIds.length, 10);

  for (let i = 0; i < maxPdfs; i++) {
    if (statusCell) {
      statusCell.setValue(`⏳ Lendo PDF ${i + 1}/${maxPdfs}...`);
      SpreadsheetApp.flush();
    }
    const txt = converterPdfParaTextoOcr(listaPdfIds[i]);
    if (txt && txt.trim().length > 20) {
      acumulado.leuAlgumPdf = true;
      const parcial = extrairDiretoDoTexto(txt);

      if (parcial.cargo !== "N/A") {
        if (/^associad[oa]$/i.test(parcial.cargo)) {
          achouAssociadoFallback = true;
        } else if (acumulado.cargo === "N/A") {
          acumulado.cargo = parcial.cargo;
        }
      }
      if (acumulado.centro_de_custo === "N/A" && parcial.centro_de_custo !== "N/A") acumulado.centro_de_custo = parcial.centro_de_custo;
      if (acumulado.horario_de_trabalho === "N/A" && parcial.horario_de_trabalho !== "N/A") acumulado.horario_de_trabalho = parcial.horario_de_trabalho;
      if (acumulado.data_admissao === "N/A" && parcial.data_admissao !== "N/A") acumulado.data_admissao = parcial.data_admissao;

      if (acumulado.cargo !== "N/A" && acumulado.centro_de_custo !== "N/A" && acumulado.horario_de_trabalho !== "N/A" && acumulado.data_admissao !== "N/A") {
        break;
      }
    }
  }

  if (acumulado.cargo === "N/A" && achouAssociadoFallback) {
    acumulado.cargo = "ASSOCIADO";
  }
  return acumulado;
}

function listarPdfsDaPastaOrdenados(pasta) {
  const p1 = [], p2 = [], p3 = [], p4 = [];
  const chavesP1 = ["ficha", "registro", "historico", "histórico", "empregado", "funcional", "fre", "cadastro", "evolucao", "evolução"];
  const chavesP2 = ["admiss", "contrato", "rh", "carteira", "ctps", "trct", "rescis", "ponto", "espelho", "contestac", "inicial", "pericia", "laudo", "ppp"];
  const chavesIgnorarOuPorUltimo = ["procurac", "preposto", "substabelecimento", "guia", "custas", "comprovante", "atestado", "habite"];

  function varrer(p, nivel) {
    if (nivel > 3) return;
    const arquivos = p.getFiles();
    while (arquivos.hasNext()) {
      const arq = arquivos.next();
      const nome = arq.getName().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      const mime = arq.getMimeType();
      if (mime === MimeType.PDF || nome.endsWith(".pdf") || mime === MimeType.GOOGLE_DOCS) {
        if (chavesP1.some(k => nome.includes(k))) p1.push(arq.getId());
        else if (chavesP2.some(k => nome.includes(k))) p2.push(arq.getId());
        else if (chavesIgnorarOuPorUltimo.some(k => nome.includes(k))) p4.push(arq.getId());
        else p3.push(arq.getId());
      }
    }
    const subpastas = p.getFolders();
    while (subpastas.hasNext()) {
      varrer(subpastas.next(), nivel + 1);
    }
  }
  varrer(pasta, 0);
  return p1.concat(p2).concat(p3).concat(p4);
}

/**
 * Converts scanned PDFs to structured text using native Google Drive OCR
 */
function converterPdfParaTextoOcr(pdfId) {
  let tempDocId = null;
  try {
    const fileOriginal = DriveApp.getFileById(pdfId);
    if (fileOriginal.getMimeType() === MimeType.GOOGLE_DOCS) {
      return DocumentApp.openById(pdfId).getBody().getText();
    }
    const blob = fileOriginal.getBlob();
    if (typeof Drive.Files.insert === "function") {
      const resource = { title: "temp_ocr_" + Date.now() };
      const file = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'pt', supportsAllDrives: true });
      tempDocId = file.id;
    } else if (typeof Drive.Files.create === "function") {
      const resource = { name: "temp_ocr_" + Date.now(), mimeType: MimeType.GOOGLE_DOCS };
      const file = Drive.Files.create(resource, blob, { ocrLanguage: 'pt', supportsAllDrives: true });
      tempDocId = file.id;
    }
    const doc = DocumentApp.openById(tempDocId);
    const texto = doc.getBody().getText();
    try {
      Drive.Files.remove(tempDocId, { supportsAllDrives: true });
    } catch (eRemove) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (_) {}
    }
    return texto;
  } catch (e) {
    Logger.log("Erro no OCR do arquivo " + pdfId + ": " + e.toString());
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (_) {}
    }
    return null;
  }
}

/**
 * HEURISTIC REGEX ENGINE v9.4 (Multi-Layer ERP/SAP Document Parser)
 */
function extrairDiretoDoTexto(textoOriginal) {
  const dados = { cargo: "N/A", centro_de_custo: "N/A", horario_de_trabalho: "N/A", data_admissao: "N/A" };
  if (!textoOriginal) return dados;
  const texto = textoOriginal.replace(/\r\n/g, "\n").replace(/[ \t]+/g, " ");

  // 1. HIRE DATE (DATA DE ADMISSÃO)
  const padroesAdmissao = [
    /(?:24\s*[-–.]?\s*)?(?:data\s+de\s+admiss[aã]o|dt\.?\s*admiss[aã]o|data\s+admiss[aã]o|admiss[aã]o\s*:|admitid[oa]\s+em|admitid[oa]\s+na\s+data|data\s+in[ií]cio|in[ií]cio\s+do\s+contrato|data\s+de\s+entrada|data\s+da\s+contrata[cç][aã]o)[\s\S]{0,250}?(\b[0-3]?\d[\/\.\-][01]?\d[\/\.\-](?:19|20)\d{2}\b)/i,
    /admiss[aã]o[\s\S]{0,180}?(\b[0-3]?\d[\/\.\-][01]?\d[\/\.\-](?:19|20)\d{2}\b)/i,
    /_{5,}[\s\n]+(\b[0-3]?\d[\/\.\-][01]?\d[\/\.\-](?:19|20)\d{2}\b)/i,
    /(?:a\s+partir\s+de|iniciando\s+em|laborou\s+de|trabalhou\s+de|v[ií]nculo\s+de|per[ií]odo\s+de)[\s\S]{0,40}?(\b[0-3]?\d[\/\.\-][01]?\d[\/\.\-](?:19|20)\d{2}\b)/i,
    /(?:admiss[aã]o|admitid[oa]|iniciando)[\s\S]{0,60}?(\b[0-3]?\d\s+de\s+(?:janeiro|fevereiro|mar[cç]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de\s+(?:19|20)\d{2})/i
  ];

  for (let r of padroesAdmissao) {
    const m = texto.match(r);
    if (m && m[1]) {
      dados.data_admissao = m[1].trim();
      break;
    }
  }

  if (dados.data_admissao === "N/A") {
    const regexFaixas = /\b([0-3]\d[\/\.\-][01]\d[\/\.\-](?:19[89]\d|20[0-2]\d))\s*(?:a|até|-)\s*[0-3]\d[\/\.\-][01]\d[\/\.\-](?:19|20|99)\d{2}\b/gi;
    let matchFaixa;
    let menorTimestamp = Infinity;
    let dataMaisAntiga = null;
    while ((matchFaixa = regexFaixas.exec(texto)) !== null) {
      const dStr = matchFaixa[1];
      const partes = dStr.split(/[\/\.\-]/);
      const dt = new Date(parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10));
      if (!isNaN(dt.getTime()) && dt.getFullYear() >= 1985 && dt.getTime() < menorTimestamp) {
        menorTimestamp = dt.getTime();
        dataMaisAntiga = dStr;
      }
    }
    if (dataMaisAntiga) dados.data_admissao = dataMaisAntiga;
  }

  // 2. JOB TITLE (CARGO / FUNÇÃO)
  const padroesCargo = [
    /\b((?:OP\.?|OPER\.?|OPERADOR[A]?)\s*(?:DE\s+)?PRODU[CÇG][AÃ]O\s*(?:I{1,3}|IV|1|2|3|JR|PL|SR)?|(?:OP\.?|OPER\.?|OPERADOR[A]?)\s+(?:DE\s+)?(?:EMPILHADEIRA|LOG[IÍ]STICA|M[AÁ]QUINAS?|INJETORA|PRENSA|PONTE|PROCESSOS?|MONTAGEM|MANUFATURA|USINAGEM|PINTURA|SOLDA)\s*(?:I{1,3}|1|2|3)?|MONTADOR[A]?\s*(?:DE\s+PRODU[CÇG][AÃ]O\s*)?(?:I{1,3}|1|2|3)?|(?:AUX\.?|AUXILIAR|AJUDANTE|ASSIST\.?|ASSISTENTE)\s+(?:DE\s+)?(?:PRODU[CÇG][AÃ]O|MANUTEN[CÇ][AÃ]O|LIMPEZA|LOG[IÍ]STICA|ENFERMAGEM|ADMINISTRATIVO|QUALIDADE|ALMOXARIFADO|LABORAT[OÓ]RIO|ENGENHARIA|EXPEDI[CÇ][AÃ]O)\s*(?:I{1,3}|1|2|3)?|(?:T[EÉ]C\.?|T[EÉ]CNIC[OA])\s+(?:DE\s+|EM\s+)?[A-ZÀ-Ú\s]{3,28}|ANALISTA\s+(?:DE\s+)?[A-ZÀ-Ú\s]{3,28}|ENGENHEIR[OA]\s+[A-ZÀ-Ú\s]{3,28}|ESPECIALISTA\s+[A-ZÀ-Ú\s]{3,28}|L[IÍ]DER\s+(?:DE\s+)?[A-ZÀ-Ú\s]{3,25}|SUPERVISOR[A]?\s+(?:DE\s+)?[A-ZÀ-Ú\s]{3,25}|COORDENADOR[A]?\s+[A-ZÀ-Ú\s]{3,25}|GERENTE\s+[A-ZÀ-Ú\s]{3,25}|INSPETOR[A]?\s+(?:DE\s+)?[A-ZÀ-Ú\s]{3,25}|MONITOR[A]?\s+(?:DE\s+)?[A-ZÀ-Ú\s]{3,25}|FERRAMENTEIRO[A]?\s*(?:I{1,3}|1|2|3)?|MANUTENTOR[A]?\s*[A-ZÀ-Ú\s]{0,20}|MEC[AÂ]NIC[OA]\s*(?:DE\s+MANUTEN[CÇ][AÃ]O|INDUSTRIAL)?\s*(?:I{1,3}|1|2|3)?|ELETRICISTA\s*(?:DE\s+MANUTEN[CÇ][AÃ]O|INDUSTRIAL)?\s*(?:I{1,3}|1|2|3)?|ELETROMEC[AÂ]NIC[OA]\s*[A-ZÀ-Ú\s]{0,15}|SOLDADOR[A]?|PREPARADOR[A]?\s+(?:DE\s+)?[A-ZÀ-Ú\s]{3,20}|REBARBADOR[A]?|PRENSEIRO[A]?|PINTOR[A]?\s*[A-ZÀ-Ú\s]{0,15}|EMPILHADEIRISTA|ALMOXARIFE|CONFERENTE|JOVEM\s+APRENDIZ|APRENDIZ|ESTAGI[AÁ]RI[OA])\b/i,
    /\b[34567]\d{7}\s*[-–]?\s*(?!Associad[oa]\b)([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú0-9\s\.]{2,38})/i,
    /(?:cargo\s*\/\s*fun[cç][aã]o|cargo\s+atual|cargo\s+inicial|22\s*[-–.]?\s*cargo|(?:exerceu|exercia|ocupava|admitid[oa]\s+para)\s+(?:o\s+cargo|a\s+fun[cç][aã]o)\s+de|cargo\s*:|fun[cç][aã]o\s*:)\s*(?!\/?\s*fun[cç][aã]o|integra[cç][aã]o|associad[oa]\b)([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú0-9\s\.]{2,40})/i,
    /\b[56]\d{7}\s*[-–]?\s*(Associad[oa])\b/i
  ];

  for (let r of padroesCargo) {
    const m = texto.match(r);
    if (m && m[1]) {
      const candidato = m[1].trim();
      if (ehCargoValidoHeuristica(candidato)) {
        dados.cargo = candidato;
        break;
      }
    }
  }

  // 3. COST CENTER (CENTRO DE CUSTO)
  const padroesCC = [
    /(?:1000|\b)([MBJCERSPAFGLTUW]\s?\d{6}\b(?:[\s\-–]+[A-ZÀ-Ú0-9][^\n,;]{1,42})?)/,
    /(?:centro\s+de\s+custo|c\.?\s*custo|c\/c|setor\s*\/\s*cc|unid(?:ade)?\.?\s*org(?:anizacional)?\.?)\s*[:\-–]\s*(?!_{3,}|data|cargo|hor[aá]rio)([A-Z0-9][^\n,;.]{2,48})/i,
    /\b((?:MONTAGEM|MANUFATURA|USINAGEM|ESTAMPARIA|INJE[CÇ][AÃ]O|PINTURA|ESMALTA[CÇ][AÃ]O|LOG[IÍ]STICA|QUALIDADE|MANUTEN[CÇ][AÃ]O|ENGENHARIA|TESOURAS|CENTRAL\s+PORTAS|PU\s+LINHA|ABASTEC\.?\s*MONTAGEM|EVAPORADORES|TERMOFORMAGEM|REFRIGERA[CÇ][AÃ]O|LAVANDERIA|COC[CÇ][AÃ]O|COMPRESSORES|EMBALAGEM|EXPEDI[CÇ][AÃ]O|FERRAMENTARIA|LABORAT[OÓ]RIO)\s+(?:L\d|F\d|FMO|LINHA\s*\d|T\d|[A-Z0-9\-]{2,18}))\b/
  ];

  for (let r of padroesCC) {
    const m = texto.match(r);
    if (m && m[1]) {
      dados.centro_de_custo = m[1].trim().replace(/^([A-Z])\s+(\d{6})/, "$1$2");
      break;
    }
  }

  // 4. WORK SHIFT / SCHEDULE (HORÁRIO DE TRABALHO)
  const textoSemDatas = texto.replace(/\b[0-3]?\d[\/\.\-][01]?\d[\/\.\-](?:19|20|99)\d{2}(?:\s*(?:a|até|-)\s*[0-3]?\d[\/\.\-][01]?\d[\/\.\-](?:19|20|99)\d{2})?/gi, " ");
  const padroesHorario = [
    /\b((?:SEG|TER|QUA|QUI|SEX|SAB|SÁB|DOM|2ª|3ª|4ª|5ª|6ª)(?:[\s\-AÀE\/]+(?:SEG|TER|QUA|QUI|SEX|SAB|SÁB|DOM|2ª|3ª|4ª|5ª|6ª|FEIRA))?[\s\-:]+(?:das\s+)?(?:[01]\d|2[0-3])[:h\.]?[0-5]\d\s*(?:[aà]s|a|-|\s)\s*(?:[01]\d|2[0-3])[:h\.]?[0-5]\d(?:[\s\-\/]+(?:SEG|TER|QUA|QUI|SEX|SAB|SÁB|DOM|ALT|[012]\d[:h\.]?[0-5]\d))*)/i,
    /\b((?:[01]\d|2[0-3])[:h][0-5]\d\s*(?:[aà]s|as|a|-|\s)\s*(?:[01]\d|2[0-3])[:h][0-5]\d(?:\s*(?:SEG|TER|QUA|QUI|SEX|SAB|SÁB|DOM|ALT)[^\n,;.]{0,20})?)/i,
    /(?:descri[cç][aã]o\s+do\s+turno|hor[aá]rio|jornada|turno|escala)[^:\n]{0,35}[:\s]+(?:ser[aá]\s+de\s+)?([^\n,;]{4,65}?\b(?:[01]\d|2[0-3])[:h\.]?[0-5]\d\b[^\n,;]{0,25})/i,
    /\b((?:1[º°]|2[º°]|3[º°]|PRIMEIRO|SEGUNDO|TERCEIRO)\s+TURNO(?:\s*[-–]\s*[^\n,;.]{3,25})?|TURNO\s+(?:COMERCIAL|ADMINISTRATIVO|NORMAL))\b/i
  ];

  for (let r of padroesHorario) {
    const m = textoSemDatas.match(r);
    if (m && m[1]) {
      dados.horario_de_trabalho = m[1].trim();
      break;
    }
  }

  return limparDadosExtraidos(dados);
}

function ehCargoValidoHeuristica(str) {
  if (!str || str.length < 3) return false;
  const s = str.trim();
  const falsosPositivos = /^(?:de|da|do|dos|das|em|na|no|para|por|com|sem|que|se|ao|os|as|um|uma|seu|sua|este|esta|isso|pelo|pela)\b/i;
  const palavrasProibidas = /\b(?:reclamante|reclamada|empresa|processo|contrato|rescis[aã]o|admiss[aã]o|demiss[aã]o|afastamento|benef[ií]cio|inss|fgts|ctps|per[ií]cia|laudo|insalubridade|periculosidade|horas\s+extras|intervalo|repouso|integra[cç][aã]o|default|padr[aã]o|p[aá]gina|folha|assinatura|documento|identidade|cpf|cnpj|pis|pasep|sindicato|acordo|conven[cç][aã]o|via\s+rh)\b/i;
  if (falsosPositivos.test(s) || palavrasProibidas.test(s)) return false;
  if (!/[A-ZÀ-Úa-zà-ú]{3,}/.test(s)) return false;
  return true;
}

function limparDadosExtraidos(dados) {
  if (dados.data_admissao && dados.data_admissao !== "N/A") {
    let strData = dados.data_admissao.toString();
    const meses = { jan: "01", fev: "02", mar: "03", abr: "04", mai: "05", jun: "06", jul: "07", ago: "08", set: "09", out: "10", nov: "11", dez: "12" };
    const mExt = strData.match(/([0-3]?\d)\s+de\s+(janeiro|fevereiro|mar[cç]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de\s+((?:19|20)\d{2})/i);
    if (mExt) {
      strData = `${mExt[1].padStart(2, "0")}/${meses[mExt[2].substring(0, 3).toLowerCase()]}/${mExt[3]}`;
    }
    const apenasData = strData.match(/([0-3]?\d)[\/\.\-]([01]?\d)[\/\.\-](\d{2,4})/);
    if (apenasData) {
      const dia = apenasData[1].padStart(2, "0");
      const mes = apenasData[2].padStart(2, "0");
      let ano = apenasData[3];
      if (ano.length === 2) ano = parseInt(ano, 10) > 50 ? "19" + ano : "20" + ano;
      dados.data_admissao = `${dia}/${mes}/${ano}`;
    } else {
      dados.data_admissao = "N/A";
    }
  }

  if (dados.cargo && dados.cargo !== "N/A") {
    let c = dados.cargo
      .replace(/[\r\n]+/g, " ")
      .replace(/\bPRODUGAO\b/gi, "PRODUCAO")
      .replace(/^[\s\/\\\-–:]+/, "")
      .replace(/^(?:fun[cç][aã]o|cargo|posi[cç][aã]o)\s*[:\-–]?\s*/i, "")
      .replace(/^(?:o\s+de|a\s+de|de|da|do|como|na\s+fun[cç][aã]o\s+de|no\s+cargo\s+de)\s+/i, "")
      .replace(/^\d{5,10}\s*[-–]?\s*/, "")
      .replace(/\s*(?:descri[cç][aã]o\s+do\s+turno|descri[cç][aã]o\s+da\s+fun|hor[aá]rio\s+de\s+trabalho|centro\s+de\s+custo|sal[aá]rio|remunera[cç][aã]o|mensal|horista|podendo\s+de\s+acordo|conforme\s+necessidade|no\s+setor|na\s+unidade|recebendo|onde|cbo|\b(?:1000)?[A-Z]\d{6}\b).*$/i, "")
      .trim()
      .toUpperCase();
    dados.cargo = ehCargoValidoHeuristica(c) ? c : "N/A";
  }

  if (dados.centro_de_custo && dados.centro_de_custo !== "N/A") {
    let cc = dados.centro_de_custo
      .replace(/[\r\n]+/g, " ")
      .replace(/^1000\s*([A-Z]\d{6})/i, "$1")
      .replace(/^(?:o\s+de|de|da|do|n[º°]?)\s+/i, "")
      .replace(/\s+(?:da|do|de)\s+/i, " - ")
      .replace(/\s+\d{2}[\/\.\-]\d{2}[\/\.\-]\d{2,4}.*$/i, "")
      .replace(/\s+(?:onde|no\s+qual|exercendo|realizava|seguro\s+de\s+vida|hor[aá]rio|jornada|turno|sal[aá]rio|descri[cç][aã]o).*$/i, "")
      .replace(/[\s\-–:]+$/g, "")
      .replace(/_{2,}/g, "")
      .trim();
    const ccInvalido = /^(?:manuten[cç][aã]o|produ[cç][aã]o|montagem|log[ií]stica|qualidade|engenharia)(?:\s*[-–]?\s*(?:de|da|do))?$/i.test(cc);
    dados.centro_de_custo = (cc && cc.length >= 3 && !ccInvalido) ? cc : "N/A";
  }

  if (dados.horario_de_trabalho && dados.horario_de_trabalho !== "N/A") {
    let h = dados.horario_de_trabalho
      .replace(/[\r\n]+/g, " ")
      .replace(/\b[0-3]?\d[\/\.\-][01]?\d[\/\.\-]\d{2}:?\d{2}(?:\s*(?:a|até|-)\s*[0-3]?\d[\/\.\-][01]?\d[\/\.\-]\d{2}:?\d{2})?/gi, "")
      .replace(/^(?:ser[aá]\s+de|era\s+de|era\s+das|das|de|no\s+hor[aá]rio\s+de|cumprindo)\s+/i, "")
      .replace(/\s*(?:horas|hrs|hs\b|,?\s*com\s+(?:meia|uma|1|01|\d+)\s*hora|,?\s*com\s+intervalo|,?\s*para\s+repouso).*$/i, "")
      .replace(/(SEG|TER|QUA|QUI|SEX|SAB|SÁB|DOM)(\d{2}:?\d{2})/gi, "$1 $2")
      .replace(/\b([01]\d|2[0-3])h([0-5]\d)\b/gi, "$1:$2")
      .replace(/(?<![\/\.\-])\b([01]\d|2[0-3])([0-5]\d)\b(?![\/\.\-])/g, "$1:$2")
      .replace(/[\)\]\.,;:\-\s]+$/g, "")
      .replace(/^[\(\[\.,;:\-\s]+/g, "")
      .replace(/\s{2,}/g, " ")
      .trim()
      .toUpperCase();

    const temHoraOuTurnoValido = /\b(?:[01]\d|2[0-3]):[0-5]\d\b/.test(h) || /\b(?:1[º°]|2[º°]|3[º°]|PRIMEIRO|SEGUNDO|TERCEIRO|COMERCIAL)\s*TURNO\b/i.test(h);
    const temPalavraProibida = /terminado|rescis[aã]o|seguro|farma|quisitivo|trabalho\s*atual/i.test(h);
    dados.horario_de_trabalho = (temHoraOuTurnoValido && !temPalavraProibida) ? h : "N/A";
  }

  return dados;
}

// ================================================================== //
// SECTION 2: AUTOMATED GOOGLE CALENDAR DISPATCHER
// ================================================================== //

function agendarCompromissos() {
  const ui = SpreadsheetApp.getUi();
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Audiencias");
  if (!planilha) return ui.alert("Aba 'Audiencias' não encontrada.");

  const dados = planilha.getDataRange().getValues();
  const calendario = CalendarApp.getDefaultCalendar();
  let totalAgendados = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = i + 1;
    const [
      reclamante, processo, data, horaInicio, horaFim, tipoCompromisso,
      unidade, emailAudiencia, emailPericia, cargo, centroCusto,
      horarioTrabalho, dataAdmissao, linkLocal, agendar, status
    ] = dados[i];

    if (agendar !== true || status === "Agendado") continue;

    try {
      const dataInicio = parseDataHora(data, horaInicio);
      const dataFim = parseDataHora(data, horaFim);

      let listaEmails = [];
      if (emailAudiencia) listaEmails.push(emailAudiencia.toString().trim());
      if (eTipoPericia(tipoCompromisso) && emailPericia) {
        listaEmails.push(emailPericia.toString().trim());
      }
      const convidados = listaEmails.filter(e => e !== "").join(',');

      let descricao = `Reclamante: ${reclamante || "Não informado"}\n` +
                      `Processo: ${processo || "Não informado"}\n` +
                      `Tipo: ${tipoCompromisso || "Não informado"}\n` +
                      `Unidade: ${unidade || "Não informada"}\n` +
                      `Local/Link: ${linkLocal || "Não informado"}\n`;

      if (eTipoPericia(tipoCompromisso)) {
        descricao += `\n--- DADOS DA PERÍCIA ---\n` +
                     `Data de Admissão: ${formatarData(dataAdmissao) || "N/A"}\n\n` +
                     `[Cargo (5 Anos)]\n${cargo || "N/A"}\n\n` +
                     `[Centro de Custo (5 Anos)]\n${centroCusto || "N/A"}\n\n` +
                     `[Horário de Trabalho (Mais Atual)]\n${horarioTrabalho || "N/A"}`;
      }

      const titulo = `${tipoCompromisso || "Compromisso"}: ${reclamante} - ${processo}`;
      const opcoes = { description: descricao, guests: convidados, sendInvites: true };

      calendario.createEvent(titulo, dataInicio, dataFim, opcoes);
      planilha.getRange(linha, 16).setValue("Agendado");
      planilha.getRange(linha, 15).setValue(false);
      totalAgendados++;
    } catch (e) {
      planilha.getRange(linha, 16).setValue("Erro: Verifique datas/horas");
    }
  }
  ui.alert(`${totalAgendados} compromisso(s) agendado(s) com sucesso!`);
}

function parseDataHora(dataBase, horaBase) {
  const d = new Date(dataBase);
  const h = new Date(horaBase);
  d.setHours(h.getHours(), h.getMinutes(), 0, 0);
  return d;
}

function formatarData(dataVal) {
  if (!dataVal || dataVal === "N/A") return "N/A";
  if (dataVal instanceof Date) {
    return Utilities.formatDate(dataVal, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return dataVal.toString().trim();
}
