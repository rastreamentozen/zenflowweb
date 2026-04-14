// ====================================================================================
// AUTOMAÇÕES DE LOTE (COM LOGÍSTICA ATUALIZADA E PARSE FIPE FIX)
// ====================================================================================
function processarItemLoteWeb(cli, comando) {
  const token = autenticarHINOVA();
  if (!token) return { status: 'erro', msg: 'Falha na autenticação Hinova.' };
  
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = ss.getSheetByName(cli.abaNome);
  if (!aba) return { status: 'erro', msg: 'Aba não encontrada.' };
  
  const vb = cli.chassi || cli.placa;
  const pb = cli.chassi ? "chassi" : "placa";
  const linha = cli.linhaOriginal;
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const codMoto = ["3", "126", "115", "100", "105", "127", "116", "32", "33", "34", "35", "95", "96", "97"];
  
  try {
    if (comando === "logistica") {
      const est = String(aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1).getValue()).trim().toUpperCase();
      const celulaEstado = aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1);
      const notaEndereço = celulaEstado.getNote();
      if (notaEndereço && notaEndereço.indexOf("Cidade:") !== -1) {
        const ssTec = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
        const listaTecnicos = ssTec.getSheets()[0].getDataRange().getValues().slice(1).filter(t => t[0]); 
          
        const cidadeMatch = notaEndereço.match(/Cidade:\s*(.*)/);
        const bairroMatch = notaEndereço.match(/Bairro:\s*(.*)/);
        const cidadeCli = cidadeMatch ? cidadeMatch[1].split('\n')[0].trim() : "";
        const bairroCli = bairroMatch ? bairroMatch[1].split('\n')[0].trim() : "";

        const enderecoDestino = `${bairroCli}, ${cidadeCli} - ${est}, Brasil`;
        let melhorDistancia = Infinity;
        let melhorTecnico = null;
        let melhorTempoSeg = 0;
        let melhorTipo = "Volante";
        
        listaTecnicos.forEach(tecnico => {
          try {
            const logradouro = tecnico[1] || "";
            const numero = tecnico[2] || "";
            const bairro = tecnico[3] || "";
            const cidade = tecnico[4] || "";
            const estado = tecnico[5] || "";
            const cep = tecnico[6] || "";
            const tipo = tecnico[8] ? String(tecnico[8]).trim() : "Volante"; 
            
            const origemCompleta = `${logradouro}, ${numero} - ${bairro}, ${cidade} - ${estado}, ${cep}, Brasil`;

            const direcoes = Maps.newDirectionFinder()
              .setOrigin(origemCompleta)
              .setDestination(enderecoDestino)
              .setMode(Maps.DirectionFinder.Mode.DRIVING)
              .getDirections();

            if (direcoes.routes && direcoes.routes.length > 0) {
              const rota = direcoes.routes[0].legs[0];
              if (rota.distance.value < melhorDistancia) {
                melhorDistancia = rota.distance.value;
                melhorTecnico = tecnico[0];
                melhorTempoSeg = rota.duration.value;
                melhorTipo = tipo; 
              }
            }
          } catch (e) {}
        });
        
        if (melhorTecnico) {
          const distKm = (melhorDistancia / 1000).toFixed(1);
          const h = Math.floor(melhorTempoSeg / 3600);
          const m = Math.floor((melhorTempoSeg % 3600) / 60);
          const tempoFormatado = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
          const notaLimpa = notaEndereço.split("\n\n--- 🛰️ LOGÍSTICA ---")[0];
          celulaEstado.setNote(notaLimpa + "\n\n--- 🛰️ LOGÍSTICA ---\n" + `Atendimento: [${melhorTipo}] "${melhorTecnico}" - ${distKm} Km / ${tempoFormatado} de distância de carro`);
          return { status: 'ok', msg: 'Rota calculada' };
        }
      }
      return { status: 'ignorado', msg: 'Endereço inválido para roteirização' };
    }

    const resp = UrlFetchApp.fetch(`${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(vb)}/${pb}`, { "method": "get", "headers": { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });
    
    if (resp.getResponseCode() === 200) {
      const j = JSON.parse(resp.getContentText());
      if (j && j.length > 0) {
        const v = j[0];
        let alterado = false;

        if (comando === "motos") {
          if (codMoto.includes(String(v.codigo_tipo_veiculo))) {
            aba.getRange(linha, 1, 1, aba.getLastColumn()).setBackground("#d1fae5");
            aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).setNote("🏍️ MOTO (SGA)");
            alterado = true;
          } else {
            aba.getRange(linha, 1, 1, aba.getLastColumn()).setBackground(null);
            aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).clearNote();
          }
        } else if (comando === "inativos" && v.codigo_situacao != null) {
          const cSit = String(v.codigo_situacao);
          const cNome = aba.getRange(linha, MAPA_COLUNAS.NOME + 1);
          if (cSit === "1" || cSit === "14") {
            cNome.setFontColor("#000000").setFontWeight("normal").clearNote();
          } else {
            cNome.setFontColor("#9C27B0").setFontWeight("bold").setNote(`⚠️ Situação SGA: ${MAPA_SITUACAO_SGA[cSit] || "Desconhecida"}\nVerificado: ${dt}`);
            alterado = true;
          }
        } else if (comando === "fipe") {
          // Correção de REGEX para limpeza efetiva e parseFloat
          const fTexto = String(v.valor_fipe || "");
          const fTextoLimpo = fTexto.replace(/R\$\s?/gi, '').replace(/\./g, '').replace(',', '.');
          const f = parseFloat(fTextoLimpo) || 0;
          
          if (fTexto) aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(v.valor_fipe);
          
          if ((codMoto.includes(String(v.codigo_tipo_veiculo)) && f > 0 && f < 20000) || (!codMoto.includes(String(v.codigo_tipo_veiculo)) && f > 0 && f < 30000)) {
            aba.getRange(linha, MAPA_COLUNAS.FIPE_BAIXA + 1).setValue(true);
            alterado = true;
          } else {
            aba.getRange(linha, MAPA_COLUNAS.FIPE_BAIXA + 1).setValue(false);
            alterado = true;
          }
        } else if (comando === "completar") {
          const dadosLinha = aba.getRange(linha, 1, 1, aba.getLastColumn()).getValues()[0];
          if (!dadosLinha[MAPA_COLUNAS.NOME] && v.nome) { aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setValue(String(v.nome).trim()); alterado = true; }
          if (!dadosLinha[MAPA_COLUNAS.EMAIL] && v.email) { aba.getRange(linha, MAPA_COLUNAS.EMAIL + 1).setValue(String(v.email).toLowerCase().trim()); alterado = true; }
          if (!dadosLinha[MAPA_COLUNAS.TELEFONE] && v.telefone_celular) { aba.getRange(linha, MAPA_COLUNAS.TELEFONE + 1).setValue(`(${v.ddd_celular || ""}) ${v.telefone_celular}`); alterado = true; }
          if (!dadosLinha[MAPA_COLUNAS.FIPE] && v.valor_fipe) { aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(v.valor_fipe); alterado = true; }
        } else if (comando === "estados") {
          let est = v.estado ? String(v.estado).trim().toUpperCase() : "";
          let cid = v.cidade ? String(v.cidade).trim() : "";
          let bai = v.bairro ? String(v.bairro).trim() : "";
          
          if ((!est || est === "N/A" || !cid) && v.codigo_associado) {
            const rA = UrlFetchApp.fetch(`https://api.hinova.com.br/api/sga/v2/associado/buscar/${v.codigo_associado}/codigo`, { "method": "get", "headers": { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });
            if (rA.getResponseCode() === 200) {
              const jA = JSON.parse(rA.getContentText());
              const aD = (Array.isArray(jA) && jA.length > 0) ? jA[0] : jA;
              
              est = aD.estado ? String(aD.estado).trim().toUpperCase() : "N/A";
              cid = aD.cidade ? String(aD.cidade).trim() : "";
              bai = aD.bairro ? String(aD.bairro).trim() : "";
            }
          }

          if (!est) est = "N/A";
          aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1).setValue(est).setNote(`📍 Cidade: ${cid}\n🏘️ Bairro: ${bai}`);
          
          if (est !== "RJ" && est !== "N/A") {
            aba.getRange(linha, MAPA_COLUNAS.TECNICO_INDISPONIVEL + 1).setValue(true);
            alterado = true;
          } else {
             alterado = true;
          }
        }
        return { status: alterado ? 'ok' : 'ignorado', msg: 'Processado' };
      }
    }
    return { status: 'ignorado', msg: 'Nenhum dado na consulta' };
  } catch (e) { 
    return { status: 'erro', msg: e.message };
  }
}

function web_atualizarDadosCliente(abaNome, linha, dados) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = ss.getSheetByName(abaNome);
    if (!aba) return "❌ Erro: Aba não encontrada no banco.";

    if (dados.nome) aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setValue(String(dados.nome).toUpperCase());
    if (dados.placa) aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).setValue(String(dados.placa).toUpperCase());
    if (dados.chassi) aba.getRange(linha, MAPA_COLUNAS.CHASSI + 1).setValue(String(dados.chassi).toUpperCase());
    if (dados.email) aba.getRange(linha, MAPA_COLUNAS.EMAIL + 1).setValue(String(dados.email).toLowerCase());
    if (dados.telefone) aba.getRange(linha, MAPA_COLUNAS.TELEFONE + 1).setValue(String(dados.telefone));
    if (dados.fipe) aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(String(dados.fipe));

    return "✅ Dados cadastrais atualizados com sucesso!";
  } catch (e) {
    return "❌ Erro ao atualizar: " + e.message;
  }
}

function executarFerramentaWebGlobal(comando) {
  try {
    if (comando === "sincronizar_erros") { conciliarErrosMailerDaemon(); return "✅ Caixa do Gmail mapeada!"; }
    if (comando === "varrer_concluidos") { return varrerConcluidosGlobalWeb(); }
    return "⚠️ Comando não reconhecido.";
  } catch (e) { 
    return "❌ Erro API: " + e.message;
  }
}

function autenticarHINOVA() {
  const cache = CacheService.getScriptCache();
  const tokenCache = cache.get("HINOVA_TOKEN");
  if (tokenCache) return tokenCache;
  try {
    const options = {
      "method": "post",
      "headers": { "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO, "Content-Type": "application/json" },
      "payload": JSON.stringify({ "usuario": SGA_CONFIG.USUARIO, "senha": SGA_CONFIG.SENHA }),
      "muteHttpExceptions": true
    };
    const resp = UrlFetchApp.fetch(SGA_CONFIG.URL_AUTH, options);
    if (resp.getResponseCode() !== 200) return null;
    
    const token = JSON.parse(resp.getContentText()).token_usuario || null;
    if (token) cache.put("HINOVA_TOKEN", token, 3600);
    return token;
  } catch (e) { return null; }
}

function varrerConcluidosGlobalWeb() {
  const token = autenticarHINOVA();
  if (!token) return "❌ Falha na autenticação com a Hinova.";
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let contConcluidos = 0;
  let logSheet = ss.getSheetByName("Log Concluídos") || ss.insertSheet("Log Concluídos");
  const todosClientes = [];
  
  ["1 -", "2 -", "3 -"].forEach(nomeFrag => {
    const aba = ss.getSheets().find(s => s.getName().includes(nomeFrag));
    if (!aba) return;
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      const placa = String(dados[i][MAPA_COLUNAS.PLACA] || "").trim().toUpperCase();
      const chassi = String(dados[i][MAPA_COLUNAS.CHASSI] || "").trim().toUpperCase();
      if (placa || chassi) {
        todosClientes.push({
          abaNome: aba.getName(), linha: i + 1, placa: placa, 
          chassi: chassi,
          nome: dados[i][MAPA_COLUNAS.NOME], fipe: dados[i][MAPA_COLUNAS.FIPE],
          email: dados[i][MAPA_COLUNAS.EMAIL], telefone: dados[i][MAPA_COLUNAS.TELEFONE],
          identificador: chassi || placa, tipoBusca: chassi ? "chassi" : "placa"
        });
      }
    }
  });
  
  if (todosClientes.length === 0) return "Nenhum cliente na fila para varrer.";
  
  const requests = todosClientes.map(cli => ({
    url: `${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(cli.identificador)}/${cli.tipoBusca}`, method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true
  }));
  
  const responses = [];
  const TAMANHO_LOTE = 20; 
  
  for (let i = 0; i < requests.length; i += TAMANHO_LOTE) {
    const pedacoRequests = requests.slice(i, i + TAMANHO_LOTE);
    const respostasPedaco = UrlFetchApp.fetchAll(pedacoRequests);
    responses.push(...respostasPedaco);
    if (i + TAMANHO_LOTE < requests.length) {
      Utilities.sleep(1000);
    }
  }
  
  const linhasParaDeletar = {};
  const dadosParaLog = [];
  
  todosClientes.forEach((cli, index) => {
    try {
      if (responses[index].getResponseCode() === 200) {
        const j = JSON.parse(responses[index].getContentText());
        if (j && j.length > 0 && String(j[0].codigo_classificacao) === "1") {
          if (!linhasParaDeletar[cli.abaNome]) linhasParaDeletar[cli.abaNome] = [];
          linhasParaDeletar[cli.abaNome].push(cli.linha);
          dadosParaLog.push([dt, cli.nome, cli.placa, cli.chassi, cli.fipe, cli.email, cli.telefone, cli.abaNome]);
          contConcluidos++;
        }
      }
    } catch (e) { }
  });
  
  if (dadosParaLog.length > 0) {
    const ultimaLinhaLog = logSheet.getLastRow() || 1;
    logSheet.getRange(ultimaLinhaLog + 1, 1, dadosParaLog.length, dadosParaLog[0].length).setValues(dadosParaLog);
  }
  
  for (const abaNome in linhasParaDeletar) {
    const aba = ss.getSheetByName(abaNome);
    if (aba) {
      linhasParaDeletar[abaNome].sort((a, b) => b - a);
      linhasParaDeletar[abaNome].forEach(linha => aba.deleteRow(linha));
    }
  }
  return `✅ Varredura Completa na API!\n${contConcluidos} veículos "Concluídos" encontrados e movidos para o Log.`;
}

function atualizarMarcacaoWeb(abaNome, linha, campo, valorBooleano) {
  try {
    const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(abaNome);
    if (!aba) return "Aba não encontrada.";
    let col = campo === 'fipeBaixa' ? MAPA_COLUNAS.FIPE_BAIXA + 1 : campo === 'tecnicoIndisp' ? MAPA_COLUNAS.TECNICO_INDISPONIVEL + 1 : campo === 'respEmail' ? MAPA_COLUNAS.RESPONDEU_EMAIL + 1 : campo === 'respWhats' ? MAPA_COLUNAS.RESPONDEU_WHATS + 1 : 0;
    if (col > 0) { 
      aba.getRange(linha, col).setValue(valorBooleano);
      return "✅ Salvo na planilha!";
    }
    return "Campo inválido.";
  } catch (e) { return "❌ Erro: " + e.message; }
}

function marcarComoEnviadoWeb(clientesSelecionados, responsavel) {
  if (!clientesSelecionados || clientesSelecionados.length === 0) return "Vazio.";
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let cont = 0;
  
  clientesSelecionados.forEach(cli => {
    const aba = ss.getSheetByName(cli.abaNome);
    if (!aba) return;
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(true);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL + 1).setValue(dt);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel);
    const chaveAcao = cli.etapaNum === 1 ? "1_EMAIL" : cli.etapaNum === 2 ? "2_EMAIL" : "3_EMAIL";
    registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, chaveAcao, dt, responsavel);
    cont++;
  });
  return `✅ ${cont} marcados na planilha!`;
}

function web_obterConfiguracoes() {
  const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("⚙️ Configurações");
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  const configs = [];
  for (let i = 1; i < dados.length; i++) { 
    if (dados[i][0]) configs.push({ linhaOriginal: i + 1, chave: String(dados[i][0]), texto: dados[i][1] ? String(dados[i][1]) : "" });
  }
  return JSON.parse(JSON.stringify(configs));
}

function salvarConfiguracaoWeb(linha, novoTexto) {
  try { 
    SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("⚙️ Configurações").getRange(linha, 2).setValue(novoTexto);
    return "✅ Template atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}