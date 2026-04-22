// ====================================================================================
// MOTOR DE IMPORTAÇÃO E FILA GERAL
// ====================================================================================
function cadastrarLoteWeb(loteDeClientes) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba1 = ss.getSheets().find(s => s.getName().includes("1 -"));
    const aba2 = ss.getSheets().find(s => s.getName().includes("2 -"));
    const aba3 = ss.getSheets().find(s => s.getName().includes("3 -"));
    
    if (!aba1 || !aba2 || !aba3) return "❌ Erro: Abas de operação não encontradas.";

    const chassisNoSistema = new Set();
    const placasNoSistema = new Set();
    
    // Mapear existentes nas abas principais
    ss.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -") || s.getName().includes("4 -")).forEach(aba => {
      const dados = aba.getDataRange().getValues();
      for (let i = 1; i < dados.length; i++) {
        if (dados[i][MAPA_COLUNAS.CHASSI - 1]) chassisNoSistema.add(String(dados[i][MAPA_COLUNAS.CHASSI - 1]).trim().toUpperCase());
        if (dados[i][MAPA_COLUNAS.PLACA - 1]) placasNoSistema.add(String(dados[i][MAPA_COLUNAS.PLACA - 1]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
      }
    });

    // Mapear existentes na Auditoria
    const abasAuditoria = ss.getSheets().filter(s => {
      const nomeAba = s.getName().toLowerCase().trim();
      return nomeAba === "log concluídos" || nomeAba.includes("auditoria");
    });
    
    abasAuditoria.forEach(aba => {
      const dados = aba.getDataRange().getValues();
      if (dados.length < 2) return;
      
      const header = dados[0] || [];
      let idxPlaca = -1, idxChassi = -1;
      
      for (let c = 0; c < header.length; c++) {
        let n = String(header[c]).toLowerCase();
        if (n === "placa" || n.includes("placa")) idxPlaca = c;
        if (n === "chassi" || n.includes("chassi")) idxChassi = c;
      }
      
      if (idxPlaca === -1 && aba.getName() === "Log Concluídos") idxPlaca = 2;
      if (idxChassi === -1 && aba.getName() === "Log Concluídos") idxChassi = 3;

      for (let i = 1; i < dados.length; i++) {
        if (idxChassi !== -1 && dados[i][idxChassi]) chassisNoSistema.add(String(dados[i][idxChassi]).trim().toUpperCase());
        if (idxPlaca !== -1 && dados[i][idxPlaca]) placasNoSistema.add(String(dados[i][idxPlaca]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
        
        if (idxPlaca === -1 && idxChassi === -1) {
           for (let c = 0; c < dados[i].length; c++) {
             let val = String(dados[i][c]).trim().toUpperCase();
             if (/^[A-Z]{3}-?[0-9][A-Z0-9][0-9]{2}$/.test(val)) placasNoSistema.add(val.replace(/[^A-Z0-9]/g, ''));
             if (/^[A-HJ-NPR-Z0-9]{17}$/.test(val)) chassisNoSistema.add(val);
           }
        }
      }
    });
    
    const qtdColunasParaInserir = Math.max(aba1.getLastColumn(), 20) - 1;
    const dtHoje = new Date();
    const dtHojeStr = Utilities.formatDate(dtHoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let contInseridos = 0, contDuplicados = 0;
    const lotesPorAba = { 1: [] };
    const notasPorAba = { 1: [] }; // Suporte a Notas Nativas
    
    loteDeClientes.forEach((cliente) => {
      const chassiCli = String(cliente.chassi || "").trim().toUpperCase();
      const placaCli = String(cliente.placa || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      
      // Valida se já existe na base
      if ((chassiCli && chassisNoSistema.has(chassiCli)) || (placaCli && placasNoSistema.has(placaCli))) { 
        contDuplicados++; 
        return; 
      }

      const novaLinha = new Array(qtdColunasParaInserir).fill("");
      const novaNota = new Array(qtdColunasParaInserir).fill("");
      
      novaLinha[0] = cliente.data || dtHojeStr;
      novaLinha[MAPA_COLUNAS.NOME - 1] = String(cliente.nome || "").trim().toUpperCase();
      novaLinha[MAPA_COLUNAS.PLACA - 1] = placaCli;
      novaLinha[MAPA_COLUNAS.CHASSI - 1] = chassiCli;
      novaLinha[MAPA_COLUNAS.FIPE - 1] = String(cliente.fipe || "").trim();
      novaLinha[MAPA_COLUNAS.EMAIL - 1] = String(cliente.email || "").trim().toLowerCase();
      novaLinha[MAPA_COLUNAS.TELEFONE - 1] = String(cliente.telefone || "").trim();
      
      // Injeção Direta do Endereço Importado
      if (cliente.estado) {
         novaLinha[MAPA_COLUNAS.ESTADO - 1] = String(cliente.estado).trim().toUpperCase();
         let cid = cliente.cidade ? String(cliente.cidade).trim() : "Não informada";
         let bai = cliente.bairro ? String(cliente.bairro).trim() : "Não informado";
         novaNota[MAPA_COLUNAS.ESTADO - 1] = `📍 Cidade: ${cid}\n🏘️ Bairro: ${bai}`;
      }
      
      lotesPorAba[1].push(novaLinha);
      notasPorAba[1].push(novaNota);
      contInseridos++;
      
      if (chassiCli) chassisNoSistema.add(chassiCli);
      if (placaCli) placasNoSistema.add(placaCli);
    });
    
    const inserirNaAba = (aba, matriz, matrizNotas) => {
      if (matriz.length === 0) return;
      const nomes = aba.getRange("C1:C").getValues();
      let ultimaLinhaReal = 1;
      for (let j = nomes.length - 1; j >= 0; j--) {
        if (String(nomes[j][0]).trim() !== "") { ultimaLinhaReal = j + 1; break; }
      }
      const range = aba.getRange(ultimaLinhaReal + 1, 2, matriz.length, qtdColunasParaInserir);
      range.setValues(matriz);
      range.setNotes(matrizNotas); // Aplica as notas em massa simultaneamente
    };
    
    inserirNaAba(aba1, lotesPorAba[1], notasPorAba[1]); 

    let msg = `✅ Lote Processado com Sucesso!\n📥 ${contInseridos} roteados para a Etapa 1.`;
    if (contDuplicados > 0) msg += `\n⚠️ ${contDuplicados} ignorados (já existiam na Fila ou na Auditoria).`;
    return msg;
  } catch (e) { 
    return "❌ Erro Crítico no Motor de Lote: " + e.message;
  }
}

function web_obterFilaGeral() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abas = ss.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -"));
  const fila = [];
  const templatesDict = getTemplatesDict(ss);

  let feriadosTime = [];
  try {
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }
  } catch(e) {}
  
  const mapHist = {};
  try {
    const abasAud = ss.getSheets().filter(s => s.getName().includes("4 -") || s.getName().toLowerCase().includes("auditoria"));
    
    abasAud.forEach(abaAud => {
      const dadosAud = abaAud.getDataRange().getValues();
      if (dadosAud.length < 2) return;
      
      const headAud = dadosAud[0];
      let idxPlaca = headAud.findIndex(c => String(c).toLowerCase().includes("placa"));
      let idxChassi = headAud.findIndex(c => String(c).toLowerCase().includes("chassi"));
      let idxDataE1 = headAud.findIndex(c => String(c).includes("1- Enviado e-mail"));
      let idxDataW1 = headAud.findIndex(c => String(c).includes("1 -Enviado whats"));
      let idxDataE2 = headAud.findIndex(c => String(c).includes("2- Enviado e-mail"));
      let idxDataW2 = headAud.findIndex(c => String(c).includes("2 -Enviado whats"));
      let idxDataE3 = headAud.findIndex(c => String(c).includes("3- Enviado e-mail"));
      let idxDataW3 = headAud.findIndex(c => String(c).includes("3 -Enviado whats"));

      for (let i = 1; i < dadosAud.length; i++) {
        let l = dadosAud[i];
        let placaStr = idxPlaca > -1 ? String(l[idxPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, '') : "";
        let chassiStr = idxChassi > -1 ? String(l[idxChassi]).toUpperCase().trim() : "";
        
        let key = placaStr || chassiStr;
        if (!key) continue;

        if (!mapHist[key]) mapHist[key] = { e1: "", e2: "", e3: "" };

        // Etapa 1
        let dE1 = idxDataE1 > -1 ? web_formatarDataSegura(l[idxDataE1]) : "";
        let dW1 = idxDataW1 > -1 ? web_formatarDataSegura(l[idxDataW1]) : "";
        if (dE1 && dE1 !== "Aguardando...") mapHist[key].e1 = dE1;
        else if (dW1 && dW1 !== "Aguardando...") mapHist[key].e1 = dW1;

        // Etapa 2
        let dE2 = idxDataE2 > -1 ? web_formatarDataSegura(l[idxDataE2]) : "";
        let dW2 = idxDataW2 > -1 ? web_formatarDataSegura(l[idxDataW2]) : "";
        if (dE2 && dE2 !== "Aguardando...") mapHist[key].e2 = dE2;
        else if (dW2 && dW2 !== "Aguardando...") mapHist[key].e2 = dW2;

        // Etapa 3
        let dE3 = idxDataE3 > -1 ? web_formatarDataSegura(l[idxDataE3]) : "";
        let dW3 = idxDataW3 > -1 ? web_formatarDataSegura(l[idxDataW3]) : "";
        if (dE3 && dE3 !== "Aguardando...") mapHist[key].e3 = dE3;
        else if (dW3 && dW3 !== "Aguardando...") mapHist[key].e3 = dW3;
      }
    });
  } catch(e) { console.log("Erro ao mapear auditoria: " + e.message); }

  const dtHoje = new Date();

  abas.forEach(aba => {
    const nomeAba = aba.getName();
    let numEtapa = nomeAba.includes("2 -") ? 2 : nomeAba.includes("3 -") ? 3 : 1;
    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaColuna < 1) return;

    const range = aba.getRange(1, 1, ultimaLinha, ultimaColuna);
    const dados = range.getValues();
    const notas = range.getNotes();

    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      
      const nome = l[MAPA_COLUNAS.NOME] ? String(l[MAPA_COLUNAS.NOME]).trim() : "";
      const placa = l[MAPA_COLUNAS.PLACA] ? String(l[MAPA_COLUNAS.PLACA]).trim() : "";
      const chassi = l[MAPA_COLUNAS.CHASSI] ? String(l[MAPA_COLUNAS.CHASSI]).trim() : "";
      
      if (!placa && !chassi && !nome) continue;

      const notaNome = (MAPA_COLUNAS.NOME < ultimaColuna && notas[i][MAPA_COLUNAS.NOME]) ? String(notas[i][MAPA_COLUNAS.NOME]) : "";
      const notaPlaca = (MAPA_COLUNAS.PLACA < ultimaColuna && notas[i][MAPA_COLUNAS.PLACA]) ? String(notas[i][MAPA_COLUNAS.PLACA]).toUpperCase() : "";
      const notaEmail = (MAPA_COLUNAS.EMAIL < ultimaColuna && notas[i][MAPA_COLUNAS.EMAIL]) ? String(notas[i][MAPA_COLUNAS.EMAIL]) : "";
      const notaEstado = (MAPA_COLUNAS.ESTADO < ultimaColuna && notas[i][MAPA_COLUNAS.ESTADO]) ? String(notas[i][MAPA_COLUNAS.ESTADO]) : "";

      let cidade = "", bairro = "";
      let tecnicoDisp = "", tecnicoDist = "", tecnicoTempo = "", tecnicoTipo = "Volante";
      
      if (notaEstado.includes("Cidade:")) {
        const parts = notaEstado.split("\n");
        cidade = parts[0] ? parts[0].replace("📍 Cidade:", "").trim() : "";
        bairro = parts[1] ? parts[1].replace("🏘️ Bairro:", "").trim() : "";
      }

      if (notaEstado.includes("🛰️ LOGÍSTICA")) {
        let logMatchNovo = notaEstado.match(/Atendimento: \[(.*?)\] "(.*?)" - (.*?) \/ (.*?) de distância/);
        if (logMatchNovo) {
          tecnicoTipo = logMatchNovo[1];
          tecnicoDisp = logMatchNovo[2];
          tecnicoDist = logMatchNovo[3];
          tecnicoTempo = logMatchNovo[4];
        } else {
          let logMatch = notaEstado.match(/Técnico Disponível: "(.*?)" - (.*?) \/ (.*?) de distância/);
          if (logMatch) {
            tecnicoDisp = logMatch[1];
            tecnicoDist = logMatch[2];
            tecnicoTempo = logMatch[3];
          }
        }
      }

      let dPlanilha = null;
      let dataEntradaStr = "Data não registrada";
      if (l[MAPA_COLUNAS.DATA] instanceof Date) {
         dPlanilha = l[MAPA_COLUNAS.DATA];
         dataEntradaStr = Utilities.formatDate(dPlanilha, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
         const strData = String(l[MAPA_COLUNAS.DATA] || "").split(" ")[0];
         if (strData && strData.includes("/")) {
            dataEntradaStr = strData;
            const partes = strData.split("/");
            if (partes.length === 3) dPlanilha = new Date(partes[2], partes[1] - 1, partes[0]);
         }
      }

      let dEmail = null;
      const valDataEmail = l[MAPA_COLUNAS.DATA_EMAIL];
      if (valDataEmail instanceof Date) {
         dEmail = valDataEmail;
      } else {
         const strData = String(valDataEmail || "").split(" ")[0];
         if (strData && strData !== "Aguardando..." && strData.includes("/")) {
             const partes = strData.split("/");
             if (partes.length === 3) dEmail = new Date(partes[2], partes[1] - 1, partes[0]);
         }
      }

      let diasDecorridosParaSLA = 0;
      let limiteBaseSLA = 10; 
      let diasUteisParaUI = 0; 

      if (dPlanilha && !isNaN(dPlanilha)) {
         try { diasUteisParaUI = calcularDiasUteis(dPlanilha, dtHoje, feriadosTime); } catch(e) {}
      }

      if (numEtapa === 1 || numEtapa === 2) {
          diasDecorridosParaSLA = diasUteisParaUI;
      } 
      else if (numEtapa === 3) {
          if (dPlanilha && !isNaN(dPlanilha) && dEmail && !isNaN(dEmail)) {
              try { diasDecorridosParaSLA = calcularDiasUteis(dPlanilha, dEmail, feriadosTime); } catch(e) {}
          } else {
              diasDecorridosParaSLA = 5;
          }
      }

      const telefone = l[MAPA_COLUNAS.TELEFONE] ? String(l[MAPA_COLUNAS.TELEFONE]).trim() : "";
      let msgWhats = "";
      
      if (telefone) {
        const idVeic = placa || chassi;
        const isPlural = String(idVeic).includes(",") || String(idVeic).includes(" e ");
        let chaveCorpo = numEtapa === 1 ? (l[MAPA_COLUNAS.FIPE_BAIXA] === true ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL") : numEtapa === 2 ? "LEMBRETE_5_DIAS" : "PRAZO_EXPIRADO";
        
        let txtCorpo = aplicarTemplate(templatesDict, chaveCorpo, nome || "Cliente", idVeic, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaStr);
        let disclaimer = aplicarTemplate(templatesDict, "WHATSAPP_DISCLAIMER", nome || "Cliente", idVeic, false, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaStr);
        msgWhats = (disclaimer && !disclaimer.includes("⚠️")) ? disclaimer + "\n\n" + txtCorpo : txtCorpo;
      }

      let etapaSugerida = numEtapa;
      if (numEtapa === 1) {
          if (diasUteisParaUI >= 5) etapaSugerida = 2;
      } else if (numEtapa === 2) {
          if (dEmail && !isNaN(dEmail)) {
              let diasPosEmail = 0;
              try { diasPosEmail = calcularDiasUteis(dEmail, dtHoje, feriadosTime); } catch(e) {}
              if (diasPosEmail >= 5) etapaSugerida = 3;
          }
      }

      const isEnviadoBol = (l[MAPA_COLUNAS.CHECK_EMAIL] === true || l[MAPA_COLUNAS.CHECK_EMAIL] === "TRUE" || l[MAPA_COLUNAS.CHECK_EMAIL] === 1);
      const isWhatsEnviadoBol = (l[MAPA_COLUNAS.CHECK_WHATS] === true || l[MAPA_COLUNAS.CHECK_WHATS] === "TRUE" || l[MAPA_COLUNAS.CHECK_WHATS] === 1);
      
      const keyBusca = placa.replace(/[^A-Z0-9]/g, '') || chassi.toUpperCase();
      const hist = mapHist[keyBusca] ? { ...mapHist[keyBusca] } : { e1: "", e2: "", e3: "" };

      let dataCurrentStr = "";
      if (dEmail && !isNaN(dEmail)) dataCurrentStr = Utilities.formatDate(dEmail, Session.getScriptTimeZone(), "dd/MM/yyyy");
      else if (valDataEmail && valDataEmail !== "Aguardando...") dataCurrentStr = String(valDataEmail).split(" ")[0];

      if (numEtapa === 1 && (isEnviadoBol || isWhatsEnviadoBol) && !hist.e1) hist.e1 = dataCurrentStr;
      if (numEtapa === 2 && (isEnviadoBol || isWhatsEnviadoBol) && !hist.e2) hist.e2 = dataCurrentStr;
      if (numEtapa === 3 && (isEnviadoBol || isWhatsEnviadoBol) && !hist.e3) hist.e3 = dataCurrentStr;

      fila.push({
        idUnico: nomeAba + "-" + (i + 1), 
        etapaNum: numEtapa, linhaOriginal: i + 1, abaNome: nomeAba, nome: nome, placa: placa, chassi: chassi,
        fipe: l[MAPA_COLUNAS.FIPE] ? String(l[MAPA_COLUNAS.FIPE]).trim() : "",
        email: l[MAPA_COLUNAS.EMAIL] ? String(l[MAPA_COLUNAS.EMAIL]).trim() : "",
        telefone: telefone,
        estado: l[MAPA_COLUNAS.ESTADO] ? String(l[MAPA_COLUNAS.ESTADO]).trim() : "",
        cidade: cidade, bairro: bairro,
        tecnicoDisp: tecnicoDisp, tecnicoDist: tecnicoDist, tecnicoTempo: tecnicoTempo, 
        tecnicoTipo: tecnicoTipo, 
        dataPlanilha: dataEntradaStr,
        dataEmail: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_EMAIL]),
        dataWhats: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_WHATS]),
        isEnviado: isEnviadoBol,
        isWhatsEnviado: isWhatsEnviadoBol,
        isRespondeuEmail: (l[MAPA_COLUNAS.RESPONDEU_EMAIL] === true || l[MAPA_COLUNAS.RESPONDEU_EMAIL] === "TRUE" || l[MAPA_COLUNAS.RESPONDEU_EMAIL] === 1),
        isRespondeuWhats: (l[MAPA_COLUNAS.RESPONDEU_WHATS] === true || l[MAPA_COLUNAS.RESPONDEU_WHATS] === "TRUE" || l[MAPA_COLUNAS.RESPONDEU_WHATS] === 1),
        isFipeBaixa: (l[MAPA_COLUNAS.FIPE_BAIXA] === true || l[MAPA_COLUNAS.FIPE_BAIXA] === "TRUE" || l[MAPA_COLUNAS.FIPE_BAIXA] === 1),
        isTecnicoIndisp: (l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === true || l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === "TRUE" || l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === 1),
        isMoto: notaPlaca.includes("MOTO"),
        isInativo: notaNome.includes("Situação SGA"),
        isErroEmail: notaEmail.includes("Erro:"),
        notaNome: notaNome, notaEmail: notaEmail, mensagemWhatsApp: msgWhats,
        diasUteisSLA: diasUteisParaUI, etapaSugerida: etapaSugerida,
        histE1: hist.e1, histE2: hist.e2, histE3: hist.e3 
      });
    }
  });
  return fila;
}

function web_migrarClientesEtapa(movimentacoes) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const abasInfo = {
      1: ss.getSheets().find(s => s.getName().includes("1 -")),
      2: ss.getSheets().find(s => s.getName().includes("2 -")),
      3: ss.getSheets().find(s => s.getName().includes("3 -"))
    };

    if (!abasInfo[1] || !abasInfo[2] || !abasInfo[3]) return "❌ Abas de operação não encontradas.";

    const operacoesOrigem = {}; 
    let sucesso = 0;

    movimentacoes.forEach(mov => {
      const partes = mov.idUnico.lastIndexOf('-');
      const abaNome = mov.idUnico.substring(0, partes);
      const linha = parseInt(mov.idUnico.substring(partes + 1));
      
      if (!operacoesOrigem[abaNome]) operacoesOrigem[abaNome] = [];
      
      const abaOrigem = ss.getSheetByName(abaNome);
      if (abaOrigem) {
        const rowData = abaOrigem.getRange(linha, 1, 1, abaOrigem.getLastColumn()).getValues()[0];
        const rowNotes = abaOrigem.getRange(linha, 1, 1, abaOrigem.getLastColumn()).getNotes()[0];
        const rowColors = abaOrigem.getRange(linha, 1, 1, abaOrigem.getLastColumn()).getFontColors()[0];
        const rowBGs = abaOrigem.getRange(linha, 1, 1, abaOrigem.getLastColumn()).getBackgrounds()[0];
        const rowWeights = abaOrigem.getRange(linha, 1, 1, abaOrigem.getLastColumn()).getFontWeights()[0];
        
        operacoesOrigem[abaNome].push({ linha: linha, novaEtapa: mov.novaEtapa, rowData, rowNotes, rowColors, rowBGs, rowWeights });
      }
    });

    for (const [abaNome, tarefas] of Object.entries(operacoesOrigem)) {
      const abaOrigem = ss.getSheetByName(abaNome);
      tarefas.sort((a, b) => b.linha - a.linha);

      tarefas.forEach(t => {
        const abaDestino = abasInfo[t.novaEtapa];
        if (abaDestino && abaOrigem.getName() !== abaDestino.getName()) {
          const lastRow = abaDestino.getLastRow() + 1;
          const rangeDestino = abaDestino.getRange(lastRow, 1, 1, t.rowData.length);
          
          rangeDestino.setValues([t.rowData]);
          rangeDestino.setNotes([t.rowNotes]);
          rangeDestino.setFontColors([t.rowColors]);
          rangeDestino.setBackgrounds([t.rowBGs]);
          rangeDestino.setFontWeights([t.rowWeights]);
          
          abaOrigem.deleteRow(t.linha);
          sucesso++;
        }
      });
    }

    return `✅ Migração efetuada! ${sucesso} clientes foram remanejados de etapa.`;
  } catch(e) {
    return "❌ Erro ao realizar a migração transacional: " + e.message;
  }
}

function gatilho_migracaoAutomaticaSLA() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abasInfo = {
    1: ss.getSheets().find(s => s.getName().includes("1 -")),
    2: ss.getSheets().find(s => s.getName().includes("2 -")),
    3: ss.getSheets().find(s => s.getName().includes("3 -"))
  };
  
  let feriadosTime = [];
  try {
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }
  } catch(e) {}

  const dtHoje = new Date();
  const movimentacoes = [];

  [1, 2].forEach(numEtapa => {
    const aba = abasInfo[numEtapa];
    if (!aba) return;
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const placa = String(l[MAPA_COLUNAS.PLACA] || "").trim();
      const chassi = String(l[MAPA_COLUNAS.CHASSI] || "").trim();
      if (!placa && !chassi) continue;

      let dBaseSLA = null;

      if (numEtapa === 1) {
         if (l[MAPA_COLUNAS.DATA] instanceof Date) {
            dBaseSLA = l[MAPA_COLUNAS.DATA];
         } else {
            const strData = String(l[MAPA_COLUNAS.DATA] || "").split(" ")[0];
            if (strData && strData.includes("/")) {
                const partes = strData.split("/");
                if (partes.length === 3) dBaseSLA = new Date(partes[2], partes[1] - 1, partes[0]);
            }
         }
      } else if (numEtapa === 2) {
         const valDataEmail = l[MAPA_COLUNAS.DATA_EMAIL];
         if (valDataEmail instanceof Date) {
            dBaseSLA = valDataEmail;
         } else {
            const strData = String(valDataEmail || "").split(" ")[0];
            if (strData && strData !== "Aguardando..." && strData.includes("/")) {
                const partes = strData.split("/");
                if (partes.length === 3) dBaseSLA = new Date(partes[2], partes[1] - 1, partes[0]);
            }
         }
      }

      if (dBaseSLA && !isNaN(dBaseSLA)) {
         try {
             const diasUteis = calcularDiasUteis(dBaseSLA, dtHoje, feriadosTime);
             if (diasUteis >= 5) {
                 const idUnico = aba.getName() + "-" + (i + 1);
                 movimentacoes.push({ idUnico: idUnico, novaEtapa: numEtapa + 1 });
             }
         } catch(e) {}
      }
    }
  });

  if (movimentacoes.length > 0) {
    web_migrarClientesEtapa(movimentacoes);
    return `✅ Migração Automática de SLA concluída! ${movimentacoes.length} clientes movidos de etapa devido à expiração dos 5 dias úteis.`;
  }
  return "✅ Varredura concluída. Nenhum cliente com SLA expirado hoje.";
}