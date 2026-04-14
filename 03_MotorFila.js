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
    
    let feriadosTime = [];
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }

    const chassisNoSistema = new Set();
    const placasNoSistema = new Set();
    
    // 1. Varredura de Duplicidades - Filas Ativas
    ss.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -") || s.getName().includes("4 -")).forEach(aba => {
      const dados = aba.getDataRange().getValues();
      for (let i = 1; i < dados.length; i++) {
        if (dados[i][MAPA_COLUNAS.CHASSI - 1]) chassisNoSistema.add(String(dados[i][MAPA_COLUNAS.CHASSI - 1]).trim().toUpperCase());
        if (dados[i][MAPA_COLUNAS.PLACA - 1]) placasNoSistema.add(String(dados[i][MAPA_COLUNAS.PLACA - 1]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
      }
    });

    // 2. Varredura de Duplicidades - Auditoria e Logs (Evita re-importar quem já passou pela esteira)
    const abasAuditoria = ss.getSheets().filter(s => {
      const nomeAba = s.getName().toLowerCase().trim();
      return nomeAba === "log concluídos" || nomeAba.includes("auditoria");
    });
    
    abasAuditoria.forEach(aba => {
      const dados = aba.getDataRange().getValues();
      if (dados.length < 2) return;
      
      const header = dados[0] || [];
      let idxPlaca = -1, idxChassi = -1;
      
      // Tenta localizar dinamicamente as colunas no Log/Auditoria
      for (let c = 0; c < header.length; c++) {
        let n = String(header[c]).toLowerCase();
        if (n === "placa" || n.includes("placa")) idxPlaca = c;
        if (n === "chassi" || n.includes("chassi")) idxChassi = c;
      }
      
      // Fallback para a aba "Log Concluídos" padrão do sistema
      if (idxPlaca === -1 && aba.getName() === "Log Concluídos") idxPlaca = 2;
      if (idxChassi === -1 && aba.getName() === "Log Concluídos") idxChassi = 3;

      for (let i = 1; i < dados.length; i++) {
        if (idxChassi !== -1 && dados[i][idxChassi]) chassisNoSistema.add(String(dados[i][idxChassi]).trim().toUpperCase());
        if (idxPlaca !== -1 && dados[i][idxPlaca]) placasNoSistema.add(String(dados[i][idxPlaca]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
        
        // Varredura Genérica (Regex) caso os cabeçalhos não batam
        if (idxPlaca === -1 && idxChassi === -1) {
           for (let c = 0; c < dados[i].length; c++) {
             let val = String(dados[i][c]).trim().toUpperCase();
             if (/^[A-Z]{3}-?[0-9][A-Z0-9][0-9]{2}$/.test(val)) placasNoSistema.add(val.replace(/[^A-Z0-9]/g, ''));
             if (/^[A-HJ-NPR-Z0-9]{17}$/.test(val)) chassisNoSistema.add(val);
           }
        }
      }
    });
    
    // Autenticação com a SGA Hinova
    const token = autenticarHINOVA();
    if (!token) return "❌ Erro: Falha na autenticação com a Hinova.";
    
    // Failsafe: Garante a formatação exata da URL Base
    const baseUrl = SGA_CONFIG.URL_CONSULTA_BASE.endsWith('/') ? SGA_CONFIG.URL_CONSULTA_BASE : SGA_CONFIG.URL_CONSULTA_BASE + '/';
    
    // Preparação de Requisições em Lote
    const requests = loteDeClientes.map(cli => {
      const vb = cli.chassi || cli.placa;
      const pb = cli.chassi ? "chassi" : "placa";
      return { url: `${baseUrl}${encodeURIComponent(vb)}/${pb}`, method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true };
    });
    
    const responses = [];
    const TAMANHO_LOTE = 20;
    
    for (let i = 0; i < requests.length; i += TAMANHO_LOTE) {
      const pedacoRequests = requests.slice(i, i + TAMANHO_LOTE);
      const respostasPedaco = UrlFetchApp.fetchAll(pedacoRequests);
      
      responses.push(...respostasPedaco);
      
      if (i + TAMANHO_LOTE < requests.length) {
        Utilities.sleep(1000); // Throttling preventivo do Google
      }
    }
    
    const qtdColunasParaInserir = Math.max(aba1.getLastColumn(), 20) - 1;
    const dtHoje = new Date();
    const dtHojeStr = Utilities.formatDate(dtHoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let contInseridos = 0, contDuplicados = 0, contIgnoradosStatus = 0;
    const lotesPorAba = { 1: [], 2: [], 3: [] };
    
    loteDeClientes.forEach((cliente, index) => {
      const chassiCli = String(cliente.chassi || "").trim().toUpperCase();
      const placaCli = String(cliente.placa || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      
      // Gateway 1: Bloqueia duplicidades ativas e auditoria
      if ((chassiCli && chassisNoSistema.has(chassiCli)) || (placaCli && placasNoSistema.has(placaCli))) { contDuplicados++; return; }

      let isCod14 = false;
      try {
        if (responses[index].getResponseCode() === 200) {
          const j = JSON.parse(responses[index].getContentText());
          
          // Tratamento Histórico: A API da Hinova pode devolver um array com o histórico do veículo.
          const arrayDados = Array.isArray(j) ? j : [j]; 
          
          // [SÊNIOR FIX]: Procura em todo o array se EXISTE pelo menos um registro como Pendente (14).
          const veiculoValido = arrayDados.find(v => v && String(v.codigo_classificacao).trim() === "14");
          
          if (veiculoValido) {
             isCod14 = true;
          }
        }
      } catch (e) { }

      // Gateway 2: Rigorosamente apenas permite entrada se encontrou Código 14
      if (!isCod14) { contIgnoradosStatus++; return; }

      let dataCriacao = web_extrairDataParaCalculo(cliente.data) || dtHoje;
      let diasUteis = calcularDiasUteis(dataCriacao, dtHoje, feriadosTime);
      let etapaAlvo = (diasUteis >= 10) ? 3 : (diasUteis >= 5) ? 2 : 1;

      const novaLinha = new Array(qtdColunasParaInserir).fill("");
      novaLinha[0] = cliente.data || dtHojeStr;
      novaLinha[MAPA_COLUNAS.NOME - 1] = String(cliente.nome || "").trim().toUpperCase();
      novaLinha[MAPA_COLUNAS.PLACA - 1] = placaCli;
      novaLinha[MAPA_COLUNAS.CHASSI - 1] = chassiCli;
      novaLinha[MAPA_COLUNAS.FIPE - 1] = String(cliente.fipe || "").trim();
      novaLinha[MAPA_COLUNAS.EMAIL - 1] = String(cliente.email || "").trim().toLowerCase();
      novaLinha[MAPA_COLUNAS.TELEFONE - 1] = String(cliente.telefone || "").trim();
      
      lotesPorAba[etapaAlvo].push(novaLinha);
      contInseridos++;
      
      // Alimenta o Set na memória para não duplicar placas idênticas dentro do mesmo "CTRL+V"
      if (chassiCli) chassisNoSistema.add(chassiCli);
      if (placaCli) placasNoSistema.add(placaCli);
    });
    
    const inserirNaAba = (aba, matriz) => {
      if (matriz.length === 0) return;
      const nomes = aba.getRange("C1:C").getValues();
      let ultimaLinhaReal = 1;
      for (let j = nomes.length - 1; j >= 0; j--) {
        if (String(nomes[j][0]).trim() !== "") { ultimaLinhaReal = j + 1; break; }
      }
      aba.getRange(ultimaLinhaReal + 1, 2, matriz.length, qtdColunasParaInserir).setValues(matriz);
    };
    
    inserirNaAba(aba1, lotesPorAba[1]); 
    inserirNaAba(aba2, lotesPorAba[2]); 
    inserirNaAba(aba3, lotesPorAba[3]);

    let msg = `✅ Lote Processado com Sucesso!\n📥 ${contInseridos} roteados para a fila.`;
    if (contDuplicados > 0) msg += `\n⚠️ ${contDuplicados} já existiam na Fila ou na Auditoria.`;
    if (contIgnoradosStatus > 0) msg += `\n🚫 ${contIgnoradosStatus} barrados (Fora do Cód. 14 SGA).`;
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
      
      // AQUI ESTAVA O PROBLEMA DE VERSIONAMENTO QUE EMBARALHOU TUDO!
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

      const telefone = l[MAPA_COLUNAS.TELEFONE] ? String(l[MAPA_COLUNAS.TELEFONE]).trim() : "";
      let msgWhats = "";
      
      if (telefone) {
        const idVeic = placa || chassi;
        const isPlural = String(idVeic).includes(",") || String(idVeic).includes(" e ");
        let chaveCorpo = numEtapa === 1 ? (l[MAPA_COLUNAS.FIPE_BAIXA] === true ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL") : numEtapa === 2 ? "LEMBRETE_5_DIAS" : "PRAZO_EXPIRADO";
        let txtCorpo = aplicarTemplate(templatesDict, chaveCorpo, nome || "Cliente", idVeic, isPlural);
        let disclaimer = aplicarTemplate(templatesDict, "WHATSAPP_DISCLAIMER", nome || "Cliente", idVeic, false);
        msgWhats = (disclaimer && !disclaimer.includes("⚠️")) ? disclaimer + "\n\n" + txtCorpo : txtCorpo;
      }

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
        dataPlanilha: (l[MAPA_COLUNAS.DATA] instanceof Date) ? Utilities.formatDate(l[MAPA_COLUNAS.DATA], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(l[MAPA_COLUNAS.DATA] || "").split(" ")[0],
        
        dataEmail: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_EMAIL]),
        dataWhats: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_WHATS]),
        isEnviado: (l[MAPA_COLUNAS.CHECK_EMAIL] === true || l[MAPA_COLUNAS.CHECK_EMAIL] === "TRUE" || l[MAPA_COLUNAS.CHECK_EMAIL] === 1),
        isWhatsEnviado: (l[MAPA_COLUNAS.CHECK_WHATS] === true || l[MAPA_COLUNAS.CHECK_WHATS] === "TRUE" || l[MAPA_COLUNAS.CHECK_WHATS] === 1),
        isRespondeuEmail: (l[MAPA_COLUNAS.RESPONDEU_EMAIL] === true || l[MAPA_COLUNAS.RESPONDEU_EMAIL] === "TRUE" || l[MAPA_COLUNAS.RESPONDEU_EMAIL] === 1),
        isRespondeuWhats: (l[MAPA_COLUNAS.RESPONDEU_WHATS] === true || l[MAPA_COLUNAS.RESPONDEU_WHATS] === "TRUE" || l[MAPA_COLUNAS.RESPONDEU_WHATS] === 1),
        isFipeBaixa: (l[MAPA_COLUNAS.FIPE_BAIXA] === true || l[MAPA_COLUNAS.FIPE_BAIXA] === "TRUE" || l[MAPA_COLUNAS.FIPE_BAIXA] === 1),
        isTecnicoIndisp: (l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === true || l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === "TRUE" || l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === 1),
        isMoto: notaPlaca.includes("MOTO"),
        isInativo: notaNome.includes("Situação SGA"),
        isErroEmail: notaEmail.includes("Erro:"),
        notaNome: notaNome, notaEmail: notaEmail, mensagemWhatsApp: msgWhats
      });
    }
  });
  return fila;
}