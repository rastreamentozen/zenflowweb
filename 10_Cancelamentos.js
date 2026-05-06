// ====================================================================================
// BANCO DE DADOS DESACOPLADO: SISTEMA DE CANCELAMENTOS (LÓGICAS)
// ====================================================================================

/**
 * Função de Auto-Healing: Cria a estrutura se a planilha estiver vazia.
 */
function inicializarBancoCancelamentos() {
  const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
  
  let abaBase = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
  if (!abaBase) {
    abaBase = ssCanc.insertSheet(NOME_ABA_BASE_CANC);
    abaBase.appendRow([
      "Registro Unico", "Data", "Situação", "Nome", "Placa", "Chassi", "Fipe", 
      "E-mail", "Telefone", "E-mail Enviado", "Data de Envio", "Whatsapp", 
      "Data de envio whats", "Responsável", "Estado"
    ]);
    abaBase.getRange("A1:O1").setFontWeight("bold").setBackground("#312e81").setFontColor("white");
    abaBase.setFrozenRows(1);
  }

  let abaLog = ssCanc.getSheetByName(NOME_ABA_LOG_CANC);
  if (!abaLog) {
    abaLog = ssCanc.insertSheet(NOME_ABA_LOG_CANC);
    abaLog.appendRow(["Data e Hora", "Nome do Cliente", "E-mail", "Veículos", "Quantidade", "Responsável", "Status"]);
    abaLog.getRange("A1:G1").setFontWeight("bold").setBackground("#9f1239").setFontColor("white");
    abaLog.setFrozenRows(1);
  }
}

/**
 * Motor de Importação Independente: Recebe matriz de dados do WebApp e insere no novo BD.
 */
function web_importarLoteCancelados(loteDados) {
  try {
    if (!loteDados || loteDados.length === 0) return { sucesso: false, mensagem: "Nenhum dado recebido." };

    inicializarBancoCancelamentos();
    const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
    const abaBase = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
    
    const dadosExistentes = abaBase.getDataRange().getValues();
    const placasExistentes = new Set();
    const chassisExistentes = new Set();
    
    for (let i = 1; i < dadosExistentes.length; i++) {
      if (dadosExistentes[i][COL_DB_CANC.PLACA]) placasExistentes.add(String(dadosExistentes[i][COL_DB_CANC.PLACA]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
      if (dadosExistentes[i][COL_DB_CANC.CHASSI]) chassisExistentes.add(String(dadosExistentes[i][COL_DB_CANC.CHASSI]).trim().toUpperCase());
    }

    const linhasParaInserir = [];
    const notasParaInserir = [];
    let contImportados = 0;
    let contDuplicados = 0;
    
    const dataImportacao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    for (let i = 0; i < loteDados.length; i++) {
      const row = loteDados[i];
      const placa = String(row.placa || "").trim().toUpperCase();
      const chassi = String(row.chassi || "").trim().toUpperCase();
      const placaLimpa = placa.replace(/[^A-Z0-9]/g, '');

      if (!placa && !chassi) continue; 

      if ((chassi && chassisExistentes.has(chassi)) || (placaLimpa && placasExistentes.has(placaLimpa))) {
        contDuplicados++;
        continue;
      }

      const idGerado = "CANC-" + new Date().getTime() + "-" + i;
      
      linhasParaInserir.push([
        idGerado,
        dataImportacao,
        String(row.situacao || "CANCELADO").trim().toUpperCase(),
        String(row.nome || "").trim().toUpperCase(),
        placa,
        chassi,
        String(row.fipe || "").trim(),
        String(row.email || "").trim().toLowerCase(),
        String(row.telefone || "").trim(),
        false, "", false, "", "", // Checkboxes e Datas Vazios
        String(row.estado || "").trim().toUpperCase()
      ]);

      // Monta a nota do endereço
      const notaEndereco = `Cidade: ${String(row.cidade || "N/A").trim()}\nBairro: ${String(row.bairro || "N/A").trim()}`;
      notasParaInserir.push([notaEndereco]);

      if (chassi) chassisExistentes.add(chassi);
      if (placaLimpa) placasExistentes.add(placaLimpa);
      contImportados++;
    }

    if (linhasParaInserir.length > 0) {
      const startRow = abaBase.getLastRow() + 1;
      // Insere os dados nas 15 colunas (A a O)
      abaBase.getRange(startRow, 1, linhasParaInserir.length, 15).setValues(linhasParaInserir);
      // Insere as notas apenas na coluna de Estado (Índice 14 / Coluna O)
      abaBase.getRange(startRow, 15, notasParaInserir.length, 1).setNotes(notasParaInserir);
    }

    return { 
      sucesso: true, 
      mensagem: `${contImportados} clientes importados com sucesso pro novo Banco de Dados. ${contDuplicados > 0 ? `(${contDuplicados} duplicatas ignoradas).` : ''}` 
    };

  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

/**
 * Lê os dados do Novo Banco de Dados e os Templates da Planilha Principal
 */
function web_obterFilaCancelamento() {
  try {
    inicializarBancoCancelamentos();
    const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
    const aba = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
    const dados = aba.getDataRange().getValues();
    const notas = aba.getDataRange().getNotes();
    if (dados.length < 2) return [];

    const ssMain = SpreadsheetApp.openById(PLANILHA_ID);
    const templatesDict = getTemplatesDict(ssMain);
    
    const fila = [];

    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const n = notas[i];
      const placa = String(l[COL_DB_CANC.PLACA] || "").trim().toUpperCase();
      const chassi = String(l[COL_DB_CANC.CHASSI] || "").trim().toUpperCase();
      const nome = String(l[COL_DB_CANC.NOME] || "").trim().toUpperCase();
      const idVeic = placa || chassi;
      
      if (!placa && !chassi && !nome) continue;

      const localStr = String(l[COL_DB_CANC.ESTADO] || "");
      const localNota = String(n[COL_DB_CANC.ESTADO] || "");
      let cidade = "", bairro = "";
      if (localNota.includes("Cidade:")) {
         const parts = localNota.split("\n");
         cidade = parts[0] ? parts[0].replace("Cidade:", "").trim() : "";
         bairro = parts[1] ? parts[1].replace("Bairro:", "").trim() : "";
      }

      let telefone = String(l[COL_DB_CANC.TELEFONE] || "").trim();
      let msgWhats = "";
      if (telefone) {
        let txtCorpo = aplicarTemplate(templatesDict, "CANC_WHATSAPP", nome || "Cliente", idVeic, false);
        let disclaimer = aplicarTemplate(templatesDict, "WHATSAPP_DISCLAIMER", nome || "Cliente", idVeic, false);
        msgWhats = (disclaimer && !disclaimer.includes("⚠️")) ? disclaimer + "\n\n" + txtCorpo : txtCorpo;
      }

      fila.push({
        idUnico: String(l[COL_DB_CANC.ID_UNICO] || "CANC-" + i), 
        linhaOriginal: i + 1, 
        abaNome: NOME_ABA_BASE_CANC, 
        nome: nome, 
        estado: localStr, 
        cidade: cidade, 
        bairro: bairro,
        placa: placa, 
        chassi: chassi, 
        fipe: String(l[COL_DB_CANC.FIPE] || "").trim(), 
        email: String(l[COL_DB_CANC.EMAIL] || "").trim().toLowerCase(),
        telefone: telefone, 
        isEnviado: (l[COL_DB_CANC.EMAIL_ENV] === true || l[COL_DB_CANC.EMAIL_ENV] === "TRUE" || l[COL_DB_CANC.EMAIL_ENV] === 1),
        dataEmail: String(l[COL_DB_CANC.DATA_EMAIL] || ""), 
        isWhatsEnviado: (l[COL_DB_CANC.WHATS_ENV] === true || l[COL_DB_CANC.WHATS_ENV] === "TRUE" || l[COL_DB_CANC.WHATS_ENV] === 1),
        dataWhats: String(l[COL_DB_CANC.DATA_WHATS] || ""), 
        responsavel: String(l[COL_DB_CANC.RESP] || ""), 
        codSituacao: String(l[COL_DB_CANC.SITUACAO] || "CANCELADO"),
        etapaNum: "CANCELAMENTO", 
        mensagemWhatsApp: msgWhats
      });
    }
    return fila;
  } catch (e) {
    return { erro: "Falha ao ler Novo BD de Cancelamentos: " + e.message };
  }
}

/**
 * Automação: Completar Dados SGA, agora apontando para o NOVO banco.
 */
function web_completarItemCancelamento(idUnico) {
  try {
    const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
    const aba = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
    const dados = aba.getDataRange().getValues();
    
    let linhaReal = -1;
    for (let i = 1; i < dados.length; i++) {
       if (String(dados[i][COL_DB_CANC.ID_UNICO]) === idUnico) { linhaReal = i + 1; break; }
    }
    
    if(linhaReal === -1) return { status: 'erro', msg: 'Registro não encontrado no BD' };
    
    const placa = String(aba.getRange(linhaReal, COL_DB_CANC.PLACA + 1).getValue()).trim().replace(/[^A-Z0-9]/gi, '');
    const chassi = String(aba.getRange(linhaReal, COL_DB_CANC.CHASSI + 1).getValue()).trim().toUpperCase();
    const veic = placa || chassi;
    const pb = placa ? "placa" : "chassi";
    
    if (!veic) return { status: 'erro', msg: 'Sem veículo para buscar' };
    
    const token = autenticarHINOVA();
    if (!token) return { status: 'erro', msg: 'Falha na autenticação SGA' };
    
    const baseUrl = SGA_CONFIG.URL_CONSULTA_BASE.endsWith('/') ? SGA_CONFIG.URL_CONSULTA_BASE : SGA_CONFIG.URL_CONSULTA_BASE + '/';
    const resp = UrlFetchApp.fetch(`${baseUrl}${encodeURIComponent(veic)}/${pb}`, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
    
    if (resp.getResponseCode() === 200) {
       const j = JSON.parse(resp.getContentText());
       const arr = Array.isArray(j) ? j : [j];
       
       if (arr.length > 0 && arr[0]) {
           const d = arr[0];
           let alterado = false;
           
           if (d.valor_fipe && !aba.getRange(linhaReal, COL_DB_CANC.FIPE + 1).getValue()) {
               aba.getRange(linhaReal, COL_DB_CANC.FIPE + 1).setValue(d.valor_fipe);
               alterado = true;
           }
           
           if (d.estado && !aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).getValue()) { 
               aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).setValue(String(d.estado).toUpperCase()); 
               alterado = true; 
           }
           if (d.cidade || d.bairro) { 
               const notaAtual = aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).getNote();
               if (!notaAtual) {
                   const novaNota = `Cidade: ${d.cidade || 'N/A'}\nBairro: ${d.bairro || 'N/A'}`;
                   aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).setNote(novaNota);
                   alterado = true;
               }
           }
           
           return alterado ? { status: 'ok', msg: 'Dados completados' } : { status: 'ok', msg: 'Nada novo a adicionar' };
       }
    }
    return { status: 'erro', msg: 'Veículo não localizado no SGA' };
  } catch(e) {
    return { status: 'erro', msg: e.message };
  }
}

/**
 * Funções de Marcação (Checkboxes) no Novo Banco
 */
function web_marcarComoEnviadoCancelamento(clientesSelecionados, responsavel) {
  if (!clientesSelecionados || clientesSelecionados.length === 0) return "Vazio.";
  const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
  const aba = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  const dados = aba.getDataRange().getValues();
  let cont = 0;
  
  clientesSelecionados.forEach(cli => {
    // Busca a linha real pelo ID Único para evitar falhas se a planilha for ordenada manualmente
    for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][COL_DB_CANC.ID_UNICO]) === cli.idUnico) {
            aba.getRange(i + 1, COL_DB_CANC.EMAIL_ENV + 1).setValue(true);
            aba.getRange(i + 1, COL_DB_CANC.DATA_EMAIL + 1).setValue(dt);
            aba.getRange(i + 1, COL_DB_CANC.RESP + 1).setValue(responsavel);
            cont++;
            break;
        }
    }
  });
  return `✅ ${cont} e-mails marcados no Novo Banco!`;
}

function web_marcarWhatsCancelamentoWeb(grupos, responsavel) {
  const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
  const aba = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const dados = aba.getDataRange().getValues();

  grupos.forEach(grupo => {
    grupo.linhas.forEach(cli => {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][COL_DB_CANC.ID_UNICO]) === cli.idUnico) {
            aba.getRange(i + 1, COL_DB_CANC.WHATS_ENV + 1).setValue(true);
            aba.getRange(i + 1, COL_DB_CANC.DATA_WHATS + 1).setValue(dt);
            aba.getRange(i + 1, COL_DB_CANC.RESP + 1).setValue(responsavel);
            break;
        }
      }
    });
  });
  return "✅ WhatsApp registrado no Banco!";
}

// ====================================================================================
// MOTOR DE DISPARO E PRÉVIA (Lê Template da Principal e audita na Nova)
// ====================================================================================
function web_obterPreviewCancelamentoAgrupadoWeb(grupos) {
  try {
    const ssMain = SpreadsheetApp.openById(PLANILHA_ID);
    const templatesDict = getTemplatesDict(ssMain); 
    let textoBruto = templatesDict["CANCELAMENTO"];
    if (!textoBruto) throw new Error("Template de CANCELAMENTO não encontrado na aba de Configurações.");
    
    let disclaimerWhats = templatesDict["WHATSAPP_DISCLAIMER"] || "";

    return grupos.map(g => {
        const nomeCliente = g.nome ? String(g.nome).trim() : "Cliente";
        const listaVeiculosFormatada = g.veiculos.map(v => "• " + String(v).toUpperCase().trim()).join("\n");
        
        let textoMensagem = textoBruto.replace(/{{NOME}}/gi, nomeCliente)
                                      .replace(/{{LISTA_VEICULOS}}/gi, listaVeiculosFormatada)
                                      .replace(/{{VEICULO}}/gi, listaVeiculosFormatada);

        const assuntoEmail = g.veiculos.length > 1 ? "Confirmação de Cancelamentos - Zen Seguros" : "Confirmação de Cancelamento - Zen Seguros";
        let htmlBody = typeof formatarComoEmail === "function" ? formatarComoEmail(textoMensagem, "Aviso de Cancelamento") : textoMensagem.replace(/\n/g, "<br>"); 

        let telefoneLimpo = ""; let whatsText = "";
        if (g.telefone) {
           telefoneLimpo = String(g.telefone).replace(/\D/g, "");
           if (telefoneLimpo.length >= 10 && !telefoneLimpo.startsWith("55")) telefoneLimpo = "55" + telefoneLimpo;
           whatsText = disclaimerWhats + "\n\n" + textoMensagem;
        }

        return Object.assign({}, g, { assunto: assuntoEmail, emailHtml: htmlBody, whatsText: whatsText, telefoneLimpo: telefoneLimpo });
    });
  } catch (e) { throw new Error("Erro ao gerar prévia de Cancelamento: " + e.message); }
}

function web_processarCancelamentoEmLote(payload) {
  try {
    if (!payload || !payload.email) throw new Error("E-mail do cliente ausente.");
    if (!payload.veiculos || payload.veiculos.length === 0) throw new Error("Lista de veículos vazia.");
    
    const ssMain = SpreadsheetApp.openById(PLANILHA_ID);
    const templatesDict = getTemplatesDict(ssMain);

    const emailDestino = String(payload.email).toLowerCase().trim();
    const nomeCliente = payload.nome ? String(payload.nome).trim() : "Cliente";
    const responsavel = payload.responsavel ? String(payload.responsavel).trim() : "Sistema Web";

    let textoBruto = templatesDict["CANCELAMENTO"] || "Olá {{NOME}},\n\nVeículos cancelados:\n{{LISTA_VEICULOS}}";
    const listaVeiculosFormatada = payload.veiculos.map(v => "• " + String(v).toUpperCase().trim()).join("\n");
    let textoMensagem = textoBruto.replace(/{{NOME}}/gi, nomeCliente).replace(/{{LISTA_VEICULOS}}/gi, listaVeiculosFormatada).replace(/{{VEICULO}}/gi, listaVeiculosFormatada);
    const assuntoEmail = payload.veiculos.length > 1 ? "Confirmação de Cancelamentos - Zen Seguros" : "Confirmação de Cancelamento - Zen Seguros";

    let htmlBody = typeof formatarComoEmail === "function" ? formatarComoEmail(textoMensagem, "Aviso de Cancelamento") : textoMensagem.replace(/\n/g, "<br>"); 
    GmailApp.sendEmail(emailDestino, assuntoEmail, textoMensagem, { htmlBody: htmlBody });

    web_registrarAuditoriaCancelamentoDB(nomeCliente, emailDestino, payload.veiculos, "Enviado com Sucesso", responsavel);
    const msgBaixa = web_marcarComoEnviadoCancelamento(payload.linhas, responsavel);

    return { sucesso: true, mensagem: "E-mail enviado para " + emailDestino + ". \n" + msgBaixa };
  } catch (erro) {
    if (payload && payload.email) web_registrarAuditoriaCancelamentoDB(payload.nome || "Desconhecido", payload.email, payload.veiculos || [], "Erro: " + erro.message, payload.responsavel || "Sistema Web");
    return { sucesso: false, erro: erro.message };
  }
}

function web_registrarAuditoriaCancelamentoDB(nome, email, veiculosArray, status, responsavel) {
  inicializarBancoCancelamentos();
  const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
  const abaAuditoria = ssCanc.getSheetByName(NOME_ABA_LOG_CANC);
  const dataHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const veiculosString = veiculosArray.join(" | "); 
  abaAuditoria.appendRow([dataHora, nome, email, veiculosString, veiculosArray.length, responsavel, status]);
}

// ====================================================================================
// MOTOR DE IMPORTAÇÃO DE CANCELADOS (VALIDAÇÃO E EFETIVAÇÃO)
// ====================================================================================

/**
 * FASE 1 (TRIAGEM): Recebe os dados, valida contra o Banco e retorna apenas os novos.
 */
function web_validarLoteCancelamento(dadosParaImportar) {
  try {
    inicializarBancoCancelamentos();
    const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
    const abaBase = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
    
    const dadosExistentes = abaBase.getDataRange().getValues();
    const placasExistentes = new Set();
    const chassisExistentes = new Set();
    
    for (let i = 1; i < dadosExistentes.length; i++) {
      if (dadosExistentes[i][COL_DB_CANC.PLACA]) placasExistentes.add(String(dadosExistentes[i][COL_DB_CANC.PLACA]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
      if (dadosExistentes[i][COL_DB_CANC.CHASSI]) chassisExistentes.add(String(dadosExistentes[i][COL_DB_CANC.CHASSI]).trim().toUpperCase());
    }

    const validos = [];
    let contDuplicados = 0;

    for (let i = 0; i < dadosParaImportar.length; i++) {
      const row = dadosParaImportar[i];
      const placa = String(row.placa || "").trim().toUpperCase();
      const chassi = String(row.chassi || "").trim().toUpperCase();
      const placaLimpa = placa.replace(/[^A-Z0-9]/g, '');

      if (!placa && !chassi) continue;

      if ((chassi && chassisExistentes.has(chassi)) || (placaLimpa && placasExistentes.has(placaLimpa))) {
        contDuplicados++;
        continue;
      }

      validos.push(row);
      
      // Proteção de memória para não duplicar clientes que estão no mesmo lote copiado
      if (chassi) chassisExistentes.add(chassi);
      if (placaLimpa) placasExistentes.add(placaLimpa);
    }

    return { 
      sucesso: true, 
      validos: validos,
      mensagem: `Triagem concluída. ${validos.length} novos. ${contDuplicados} duplicatas barradas.` 
    };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * FASE 2 (EFETIVAÇÃO): Recebe apenas os clientes aprovados na Triagem e grava no BD.
 */
function web_efetivarImportacaoCancelamento(payload) {
  try {
    if (!payload || payload.length === 0) return { sucesso: false, mensagem: "Nenhum dado recebido." };

    const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
    const abaBase = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
    
    const linhasParaInserir = [];
    const notasParaInserir = [];
    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    for (let i = 0; i < payload.length; i++) {
      const row = payload[i];
      const idGerado = "CANC-" + new Date().getTime() + "-" + i;
      
      linhasParaInserir.push([
        idGerado,
        row.dataPlanilha || dataAtual, // Preserva a data que o SGA gerou, se houver
        String(row.situacao || "CANCELAMENTO IMPORTADO").trim().toUpperCase(),
        String(row.nome || "").trim().toUpperCase(),
        String(row.placa || "").trim().toUpperCase(),
        String(row.chassi || "").trim().toUpperCase(),
        String(row.fipe || "").trim(),
        String(row.email || "").trim().toLowerCase(),
        String(row.telefone || "").trim(),
        false, "", false, "", "", // Checks e datas vazios
        String(row.estado || "").trim().toUpperCase()
      ]);

      const notaEndereco = `Cidade: ${String(row.cidade || "N/A").trim()}\nBairro: ${String(row.bairro || "N/A").trim()}`;
      notasParaInserir.push([notaEndereco]);
    }

    if (linhasParaInserir.length > 0) {
      const startRow = abaBase.getLastRow() + 1;
      abaBase.getRange(startRow, 1, linhasParaInserir.length, 15).setValues(linhasParaInserir);
      abaBase.getRange(startRow, 15, notasParaInserir.length, 1).setNotes(notasParaInserir);
    }

    return { sucesso: true, mensagem: `${linhasParaInserir.length} clientes salvos no Banco de Cancelamentos!` };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

// ====================================================================================
// REVISÃO DA AUTOMAÇÃO SGA (ALINHADO COM A NOVA PLANILHA)
// ====================================================================================
function web_completarItemCancelamento(idUnico) {
  try {
    const ssCanc = SpreadsheetApp.openById(PLANILHA_CANCELAMENTOS_ID);
    const aba = ssCanc.getSheetByName(NOME_ABA_BASE_CANC);
    const dados = aba.getDataRange().getValues();
    
    let linhaReal = -1;
    for (let i = 1; i < dados.length; i++) {
       if (String(dados[i][COL_DB_CANC.ID_UNICO]) === idUnico) { linhaReal = i + 1; break; }
    }
    
    if(linhaReal === -1) return { status: 'erro', msg: 'Registro não encontrado no BD' };
    
    // Lê a Placa e o Chassi da planilha para buscar
    const placa = String(aba.getRange(linhaReal, COL_DB_CANC.PLACA + 1).getValue()).trim().replace(/[^A-Z0-9]/gi, '');
    const chassi = String(aba.getRange(linhaReal, COL_DB_CANC.CHASSI + 1).getValue()).trim().toUpperCase();
    const veic = placa || chassi;
    const pb = placa ? "placa" : "chassi";
    
    if (!veic) return { status: 'erro', msg: 'Sem veículo para buscar' };
    
    const token = autenticarHINOVA();
    if (!token) return { status: 'erro', msg: 'Falha na autenticação SGA' };
    
    const baseUrl = SGA_CONFIG.URL_CONSULTA_BASE.endsWith('/') ? SGA_CONFIG.URL_CONSULTA_BASE : SGA_CONFIG.URL_CONSULTA_BASE + '/';
    const resp = UrlFetchApp.fetch(`${baseUrl}${encodeURIComponent(veic)}/${pb}`, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
    
    if (resp.getResponseCode() === 200) {
       const j = JSON.parse(resp.getContentText());
       const arr = Array.isArray(j) ? j : [j];
       
       if (arr.length > 0 && arr[0]) {
           const d = arr[0];
           let alterado = false;
           
           // Coluna G (Índice 6) = FIPE
           if (d.valor_fipe && !aba.getRange(linhaReal, COL_DB_CANC.FIPE + 1).getValue()) {
               aba.getRange(linhaReal, COL_DB_CANC.FIPE + 1).setValue(d.valor_fipe);
               alterado = true;
           }
           
           // Coluna O (Índice 14) = ESTADO (Com nota de Cidade e Bairro)
           if (d.estado && !aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).getValue()) { 
               aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).setValue(String(d.estado).toUpperCase()); 
               alterado = true; 
           }
           if (d.cidade || d.bairro) { 
               const notaAtual = aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).getNote();
               if (!notaAtual) {
                   const novaNota = `Cidade: ${d.cidade || 'N/A'}\nBairro: ${d.bairro || 'N/A'}`;
                   aba.getRange(linhaReal, COL_DB_CANC.ESTADO + 1).setNote(novaNota);
                   alterado = true;
               }
           }
           
           return alterado ? { status: 'ok', msg: 'Dados completados' } : { status: 'ok', msg: 'Nada novo a adicionar' };
       }
    }
    return { status: 'erro', msg: 'Veículo não localizado no SGA' };
  } catch(e) {
    return { status: 'erro', msg: e.message };
  }
}