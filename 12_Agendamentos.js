// ====================================================================================
// 🧠 ARQUIVO: 12_Agendamentos.gs (Back-end: Controle Geral e Instalações)
// ====================================================================================

const MAPA_TERMOS_SGA = { 
  "INSTALAÇÃO": 10, 
  "REINSTALAÇÃO": 10, 
  "RETIRADA": 14, 
  "MANUTENÇÃO": 15 
};

/**
 * Retorna a aba do mês atual na Planilha Vigente.
 */
function obterAbaAgendamentoSegura(ss) {
  const meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  const data = new Date();
  const nomeMesAtual = `${meses[data.getMonth()]} ${data.getFullYear()}`;
  let aba = ss.getSheetByName(nomeMesAtual);
  if (!aba) aba = ss.getSheets()[0];
  return aba;
}

/**
 * 🚀 FUNÇÃO RESTAURADA: Puxa a lista de meses para o Dropdown
 */
function web_obterListaMeses() {
  const ssVigente = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  const ssHist = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
  const abasV = ssVigente.getSheets().map(s => s.getName());
  const abasH = ssHist.getSheets().map(s => s.getName());
  return [...new Set([...abasV, ...abasH])].sort().reverse();
}

/**
 * 🚀 FUNÇÃO RESTAURADA: Salva a edição do Modal
 */
function web_salvarEdicaoAgendamento(obj) {
  try {
    const ssV = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
    let aba = ssV.getSheetByName(obj.aba);
    if (!aba) {
      const ssH = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
      aba = ssH.getSheetByName(obj.aba);
    }
    const linha = parseInt(obj.linha);
    aba.getRange(linha, MAPA_AGENDAMENTO.IMEI).setValue(obj.imei);
    aba.getRange(linha, MAPA_AGENDAMENTO.TECNICO).setValue(obj.tecnico);
    aba.getRange(linha, MAPA_AGENDAMENTO.DATA_AGENDAMENTO).setValue(obj.data);
    aba.getRange(linha, MAPA_AGENDAMENTO.TURNO).setValue(obj.turno);
    return { sucesso: true };
  } catch (e) { throw new Error(e.message); }
}

/**
 * FUNÇÃO SÊNIOR: Recorta o registro da origem e move para o mês vigente se necessário.
 */
function moverParaVigenteSeNecessario(linhaOriginal, nomeAbaOrigem) {
  const ssVigente = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  const abaDestino = obterAbaAgendamentoSegura(ssVigente);
  const nomeAbaDestino = abaDestino.getName();

  // Se já estiver na aba correta do mês vigente, apenas retorna a referência
  if (nomeAbaOrigem === nomeAbaDestino) {
    return { aba: abaDestino, linha: linhaOriginal, foiMovido: false };
  }

  // Caso contrário, inicia o processo de "Recortar e Colar"
  let ssOrigem = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  let abaFonte = ssOrigem.getSheetByName(nomeAbaOrigem);
  
  if (!abaFonte) {
    ssOrigem = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
    abaFonte = ssOrigem.getSheetByName(nomeAbaOrigem);
  }

  if (!abaFonte) throw new Error("Aba de origem não encontrada para mover registro.");

  // Captura os dados da linha inteira
  const rangeOriginal = abaFonte.getRange(linhaOriginal, 1, 1, abaFonte.getLastColumn());
  const dadosLinha = rangeOriginal.getValues()[0];

  // Insere na Planilha Vigente
  abaDestino.appendRow(dadosLinha);
  const novaLinha = abaDestino.getLastRow();

  // Remove da planilha antiga (Recortar)
  abaFonte.deleteRow(linhaOriginal);

  return { aba: abaDestino, linha: novaLinha, foiMovido: true };
}

/**
 * Emite o termo e MOVE o cliente para o mês vigente se ele for de um mês passado.
 */
function web_emitirTermoAgendamento(linhaPlanilha, placaOuChassi, tipoServico, abaOrigem) {
  try {
    const tipoFormatado = String(tipoServico).trim().toUpperCase();
    let codigoTermoSGA = MAPA_TERMOS_SGA[tipoFormatado];

    if (!codigoTermoSGA) {
       if(tipoFormatado.includes("INSTALA")) codigoTermoSGA = 10;
       else if(tipoFormatado.includes("RETIRA")) codigoTermoSGA = 14;
       else if(tipoFormatado.includes("MANUTEN")) codigoTermoSGA = 15;
       else throw new Error(`O serviço '${tipoServico}' não tem termo mapeado.`);
    }

    const resultadoSGA = web_emitirTermoHinova(placaOuChassi, codigoTermoSGA);
    if (resultadoSGA.erro) throw new Error(resultadoSGA.erro);

    // Transbordo: Move para o mês vigente antes de marcar a data
    const transbordo = moverParaVigenteSeNecessario(linhaPlanilha, abaOrigem);
    
    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    transbordo.aba.getRange(transbordo.linha, MAPA_AGENDAMENTO.DATA_ENVIO_TERMO).setValue(dataAtual);

    return { 
      sucesso: true, 
      link: resultadoSGA.link, 
      dataRegistrada: dataAtual,
      movido: transbordo.foiMovido 
    };

  } catch (e) {
    return { erro: e.message };
  }
}

/**
 * Conclui o serviço e MOVE o cliente para o mês vigente se ele for de um mês passado.
 */
function web_concluirServicoInteligente(linha, abaOrigem, usuarioNome) {
  try {
    const transbordo = moverParaVigenteSeNecessario(linha, abaOrigem);
    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    transbordo.aba.getRange(transbordo.linha, MAPA_AGENDAMENTO.DATA_CONCLUSAO).setValue(dataAtual);
    transbordo.aba.getRange(transbordo.linha, MAPA_AGENDAMENTO.USUARIO_CONCLUIU).setValue(usuarioNome);
    transbordo.aba.getRange(transbordo.linha, MAPA_AGENDAMENTO.SITUACAO).setValue("Concluido");
    
    return { 
      sucesso: true, 
      movido: transbordo.foiMovido 
    };
  } catch (e) {
    return { erro: e.message };
  }
}

/**
 * Listagem de Agendamentos com Otimização de Leitura
 */
function web_obterAgendamentosPorMes(nomeAba) {
  try {
    const ssV = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
    let aba = ssV.getSheetByName(nomeAba);
    
    if (!aba) {
      const ssH = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
      aba = ssH.getSheetByName(nomeAba);
    }
    if (!aba) throw new Error(`Aba '${nomeAba}' não encontrada.`);

    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return [];

    const dados = aba.getRange(1, 1, ultimaLinha, aba.getLastColumn()).getDisplayValues();
    const resultado = [];
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (!linha[MAPA_AGENDAMENTO.NOME - 1] || linha[MAPA_AGENDAMENTO.NOME - 1].toString().trim() === "") continue;
        
        resultado.push({
            linhaPlanilha: i + 1,
            abaOrigem: nomeAba,
            servico: String(linha[MAPA_AGENDAMENTO.SERVICO - 1] || ""),
            nome: String(linha[MAPA_AGENDAMENTO.NOME - 1] || ""),
            placa: String(linha[MAPA_AGENDAMENTO.PLACA - 1] || ""),
            imei: String(linha[MAPA_AGENDAMENTO.IMEI - 1] || ""),
            dataEnvioTermo: String(linha[MAPA_AGENDAMENTO.DATA_ENVIO_TERMO - 1] || ""),
            dataConclusao: String(linha[MAPA_AGENDAMENTO.DATA_CONCLUSAO - 1] || ""),
            tecnico: String(linha[MAPA_AGENDAMENTO.TECNICO - 1] || "Não Atribuído"),
            regiao: String(linha[MAPA_AGENDAMENTO.REGIAO - 1] || ""),
            valorFipe: String(linha[MAPA_AGENDAMENTO.VALOR_FIPE - 1] || "R$ 0,00"),
            fipeTipo: String(linha[MAPA_AGENDAMENTO.FIPE_TIPO - 1] || ""),
            usuarioIncluiu: String(linha[MAPA_AGENDAMENTO.USUARIO_INCLUIU - 1] || ""),
            usuarioConcluiu: String(linha[MAPA_AGENDAMENTO.USUARIO_CONCLUIU - 1] || ""),
            situacao: String(linha[MAPA_AGENDAMENTO.SITUACAO - 1] || "Pendente"),
            dataAgendamento: String(linha[MAPA_AGENDAMENTO.DATA_AGENDAMENTO - 1] || ""),
            turno: String(linha[MAPA_AGENDAMENTO.TURNO - 1] || "")
        });
    }

    return resultado.reverse();

  } catch (e) { 
    return { erro: e.message }; 
  }
}