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
 * Puxa a lista de meses para o Dropdown
 */
function web_obterListaMeses() {
  const ssVigente = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  const ssHist = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
  const abasV = ssVigente.getSheets().map(s => s.getName());
  const abasH = ssHist.getSheets().map(s => s.getName());
  return [...new Set([...abasV, ...abasH])].sort().reverse();
}

/**
 * Salva a edição do Modal
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
    
    aba.getRange(linha, MAPA_AGENDAMENTO.SERVICO).setValue(obj.servico);
    aba.getRange(linha, MAPA_AGENDAMENTO.SITUACAO).setValue(obj.situacao);
    
    if (MAPA_AGENDAMENTO.VOLUNTARIO) {
      aba.getRange(linha, MAPA_AGENDAMENTO.VOLUNTARIO).setValue(obj.voluntario);
    }
    
    aba.getRange(linha, MAPA_AGENDAMENTO.IMEI).setValue(obj.imei);
    aba.getRange(linha, MAPA_AGENDAMENTO.TECNICO).setValue(obj.tecnico);
    aba.getRange(linha, MAPA_AGENDAMENTO.DATA_AGENDAMENTO).setValue(obj.data);
    aba.getRange(linha, MAPA_AGENDAMENTO.TURNO).setValue(obj.turno);

    if (obj.nome && MAPA_AGENDAMENTO.NOME) aba.getRange(linha, MAPA_AGENDAMENTO.NOME).setValue(obj.nome);
    if (obj.placa && MAPA_AGENDAMENTO.PLACA) aba.getRange(linha, MAPA_AGENDAMENTO.PLACA).setValue(obj.placa);
    
    return { sucesso: true };
  } catch (e) { 
    throw new Error(e.message); 
  }
}

/**
 * Recorta o registro da origem e move para o mês vigente se necessário.
 */
function moverParaVigenteSeNecessario(linhaOriginal, nomeAbaOrigem) {
  const ssVigente = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  const abaDestino = obterAbaAgendamentoSegura(ssVigente);
  const nomeAbaDestino = abaDestino.getName();

  if (nomeAbaOrigem === nomeAbaDestino) {
    return { aba: abaDestino, linha: linhaOriginal, foiMovido: false };
  }

  let ssOrigem = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  let abaFonte = ssOrigem.getSheetByName(nomeAbaOrigem);
  
  if (!abaFonte) {
    ssOrigem = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
    abaFonte = ssOrigem.getSheetByName(nomeAbaOrigem);
  }

  if (!abaFonte) throw new Error("Aba de origem não encontrada para mover registro.");

  const rangeOriginal = abaFonte.getRange(linhaOriginal, 1, 1, abaFonte.getLastColumn());
  const dadosLinha = rangeOriginal.getValues()[0];

  abaDestino.appendRow(dadosLinha);
  const novaLinha = abaDestino.getLastRow();

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

    const numColunas = Math.max(aba.getLastColumn(), 20);
    const dados = aba.getRange(1, 1, ultimaLinha, numColunas).getDisplayValues();
    
    const resultado = []; 
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        
        const nomeStr = String(linha[MAPA_AGENDAMENTO.NOME - 1] || "").trim();
        const placaStr = String(linha[MAPA_AGENDAMENTO.PLACA - 1] || "").trim();
        const servicoStr = String(linha[MAPA_AGENDAMENTO.SERVICO - 1] || "").trim();
        
        if (nomeStr === "" && placaStr === "" && servicoStr === "") continue;
        
        resultado.push({
            linhaPlanilha: i + 1,
            abaOrigem: nomeAba,
            servico: servicoStr,
            nome: nomeStr,
            placa: placaStr,
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
            voluntario: MAPA_AGENDAMENTO.VOLUNTARIO ? String(linha[MAPA_AGENDAMENTO.VOLUNTARIO - 1] || "") : "", 
            dataAgendamento: String(linha[MAPA_AGENDAMENTO.DATA_AGENDAMENTO - 1] || ""),
            turno: String(linha[MAPA_AGENDAMENTO.TURNO - 1] || "")
        });
    }

    return resultado.reverse();

  } catch (e) { 
    return { erro: e.message }; 
  }
}

/**
 * 🚀 Adiciona um novo registro caçando o primeiro buraco disponível (Colunas B, C, D vazias)
 */
function web_salvarNovoAgendamento(obj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
    const aba = obterAbaAgendamentoSegura(ss);
    
    // 1. Puxa apenas as colunas A até D para caçar o primeiro buraco rapidamente
    const rangeBusca = aba.getRange("A:D");
    const valores = rangeBusca.getValues();
    
    let linhaParaInserir = -1;
    let idExistente = "";
    
    // 2. Loop de cima para baixo (começando da linha 2, ignorando o cabeçalho)
    for (let i = 1; i < valores.length; i++) {
        // Índices 1 = Coluna B, 2 = Coluna C, 3 = Coluna D
        let valB = String(valores[i][1] || "").trim();
        let valC = String(valores[i][2] || "").trim();
        let valD = String(valores[i][3] || "").trim();
        
        // Se a magia não detectar nada nas colunas B, C e D, este é o alvo!
        if (valB === "" && valC === "" && valD === "") {
            linhaParaInserir = i + 1; // Array começa no zero, planilha começa no 1
            idExistente = valores[i][0]; // Captura a numeração que já está gravada na Coluna A
            break;
        }
    }
    
    // Se a guarnição estiver lotada e não houver buraco, cria uma nova no fundo da planilha
    if (linhaParaInserir === -1) {
        linhaParaInserir = aba.getLastRow() + 1;
        idExistente = aba.getLastRow(); // Apenas um fallback para o ID
    }
    
    // 3. Prepara a matriz com o tamanho correto das colunas
    const maxCol = aba.getLastColumn() || 20; 
    const novaLinha = new Array(maxCol).fill("");

    // 4. Esmaga o índice 0 (Coluna A) com a numeração prévia (Preserva o ID)
    novaLinha[0] = idExistente;

    // 5. Preenche os campos herdados e configurados
    if (MAPA_AGENDAMENTO.NOME) novaLinha[MAPA_AGENDAMENTO.NOME - 1] = obj.nome;
    if (MAPA_AGENDAMENTO.PLACA) novaLinha[MAPA_AGENDAMENTO.PLACA - 1] = obj.placa;
    if (MAPA_AGENDAMENTO.SERVICO) novaLinha[MAPA_AGENDAMENTO.SERVICO - 1] = obj.servico;
    if (MAPA_AGENDAMENTO.SITUACAO) novaLinha[MAPA_AGENDAMENTO.SITUACAO - 1] = obj.situacao;
    if (MAPA_AGENDAMENTO.TECNICO) novaLinha[MAPA_AGENDAMENTO.TECNICO - 1] = obj.tecnico;
    if (MAPA_AGENDAMENTO.DATA_AGENDAMENTO) novaLinha[MAPA_AGENDAMENTO.DATA_AGENDAMENTO - 1] = obj.data;
    if (MAPA_AGENDAMENTO.TURNO) novaLinha[MAPA_AGENDAMENTO.TURNO - 1] = obj.turno;
    if (MAPA_AGENDAMENTO.USUARIO_INCLUIU) novaLinha[MAPA_AGENDAMENTO.USUARIO_INCLUIU - 1] = obj.usuario;
    
    // Mapeamento dinâmico para FIPE e Região
    if (MAPA_AGENDAMENTO.VALOR_FIPE) novaLinha[MAPA_AGENDAMENTO.VALOR_FIPE - 1] = obj.fipe;
    if (MAPA_AGENDAMENTO.REGIAO) novaLinha[MAPA_AGENDAMENTO.REGIAO - 1] = obj.regiao;
    // 6. Crava a atualização violenta na linha exata!
    aba.getRange(linhaParaInserir, 1, 1, maxCol).setValues([novaLinha]);
    
    return { sucesso: true };
  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  }
}
/**
 * 🚀 Recebe um Lote da Automação e grava de uma vez na aba Vigente
 */
/**
 * 🚀 Recebe um Lote da Automação da Fila e grava de uma vez na aba Vigente do Controle Geral
 */
function web_salvarLoteAgendamentos(loteArray) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
    
    const meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    const dtAtual = new Date();
    const nomeMesAtual = `${meses[dtAtual.getMonth()]} ${dtAtual.getFullYear()}`;
    let aba = ss.getSheetByName(nomeMesAtual);
    if (!aba) aba = ss.getSheets()[0]; 
    
    const rangeBusca = aba.getRange("A:D");
    const valores = rangeBusca.getValues();
    let linhaBase = -1;
    let idAtual = 0;
    
    for (let i = 1; i < valores.length; i++) {
        let valB = String(valores[i][1] || "").trim();
        let valC = String(valores[i][2] || "").trim();
        let valD = String(valores[i][3] || "").trim();
        if (valB === "" && valC === "" && valD === "") {
            linhaBase = i + 1; 
            idAtual = parseInt(valores[i][0]) || i; 
            break;
        }
    }
    
    if (linhaBase === -1) {
        linhaBase = aba.getLastRow() + 1;
        idAtual = aba.getLastRow();
    }
    
    const maxCol = aba.getLastColumn() || 20; 
    const matrizInsercao = [];

    // Mapeamento dos dados que chegam da exportação (Preservando FIPE Alta/Baixa)
    loteArray.forEach((obj, index) => {
        const novaLinha = new Array(maxCol).fill("");
        novaLinha[0] = idAtual + index;

        if (MAPA_AGENDAMENTO.NOME) novaLinha[MAPA_AGENDAMENTO.NOME - 1] = obj.nome;
        if (MAPA_AGENDAMENTO.PLACA) novaLinha[MAPA_AGENDAMENTO.PLACA - 1] = obj.placa;
        if (MAPA_AGENDAMENTO.SERVICO) novaLinha[MAPA_AGENDAMENTO.SERVICO - 1] = obj.servico;
        if (MAPA_AGENDAMENTO.SITUACAO) novaLinha[MAPA_AGENDAMENTO.SITUACAO - 1] = obj.situacao;
        if (MAPA_AGENDAMENTO.TECNICO) novaLinha[MAPA_AGENDAMENTO.TECNICO - 1] = obj.tecnico;
        if (MAPA_AGENDAMENTO.DATA_AGENDAMENTO) novaLinha[MAPA_AGENDAMENTO.DATA_AGENDAMENTO - 1] = obj.data;
        if (MAPA_AGENDAMENTO.TURNO) novaLinha[MAPA_AGENDAMENTO.TURNO - 1] = obj.turno;
        if (MAPA_AGENDAMENTO.USUARIO_INCLUIU) novaLinha[MAPA_AGENDAMENTO.USUARIO_INCLUIU - 1] = obj.usuario;
        if (MAPA_AGENDAMENTO.VALOR_FIPE) novaLinha[MAPA_AGENDAMENTO.VALOR_FIPE - 1] = obj.fipe;
        if (MAPA_AGENDAMENTO.FIPE_TIPO) novaLinha[MAPA_AGENDAMENTO.FIPE_TIPO - 1] = obj.fipeTipo;
        if (MAPA_AGENDAMENTO.REGIAO) novaLinha[MAPA_AGENDAMENTO.REGIAO - 1] = obj.regiao;

        matrizInsercao.push(novaLinha);
    });

    aba.getRange(linhaBase, 1, matrizInsercao.length, maxCol).setValues(matrizInsercao);
    return { sucesso: true };
  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  }
}