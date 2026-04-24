// ====================================================================================
// 🧠 ARQUIVO: 11_Documentos.gs (Integração com API de Termos SGA)
// ====================================================================================

/**
 * Função interna para autenticar no SGA e obter o Token Real de Sessão
 * Utiliza as credenciais da Hinova presentes no 00_Config.js
 */
function obterTokenHinovaAutenticado() {
  const urlAuth = SGA_CONFIG.URL_AUTH;
  
  const payload = {
    "usuario": SGA_CONFIG.USUARIO,
    "senha": SGA_CONFIG.SENHA
  };

  const opcoes = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const res = UrlFetchApp.fetch(urlAuth, opcoes);
  const status = res.getResponseCode();
  const body = res.getContentText();

  if (status !== 200) {
    throw new Error("Falha na Autenticação Hinova (Status " + status + "): " + body);
  }

  const json = JSON.parse(body);
  
  // Mapeamento ajustado para capturar a chave específica da Hinova (token_usuario)
  if (json.token_usuario) return json.token_usuario;
  if (json.token) return json.token;
  if (json.access_token) return json.access_token;
  
  throw new Error("SGA autenticou, mas não devolveu a chave de token esperada no JSON: " + body);
}

/**
 * Função chamada pelo front-end para emitir o contrato.
 */
function web_emitirTermoHinova(placaOuChassi, codigoTermoSGA) {
  try {
    
    // 1. Handshake: Obter o token de uso real validando Usuário e Senha do SGA
    const tokenSessao = obterTokenHinovaAutenticado();

    // 2. Buscar os IDs obrigatórios (Veículo e Associado)
    const urlBusca = SGA_CONFIG.URL_CONSULTA_BASE + placaOuChassi;
    const opcoesBusca = {
      "method": "get",
      "headers": { 
        "Authorization": "Bearer " + tokenSessao,
        "Content-Type": "application/json"
      },
      "muteHttpExceptions": true
    };
    
    const resBusca = UrlFetchApp.fetch(urlBusca, opcoesBusca);
    const statusBusca = resBusca.getResponseCode();
    const bodyBusca = resBusca.getContentText();

    if (statusBusca !== 200) {
      throw new Error("Falha na busca do veículo. Status SGA: " + statusBusca + " | Detalhe: " + bodyBusca);
    }
    
    let dadosVeiculo = JSON.parse(bodyBusca);
    
    // Tratamento de segurança: A Hinova pode retornar Objeto ou Array
    let veiculo = Array.isArray(dadosVeiculo) ? dadosVeiculo[0] : dadosVeiculo;
    
    if (!veiculo || !veiculo.codigo_associado || !veiculo.codigo_veiculo) {
      throw new Error("Veículo não localizado na Hinova ou sem os códigos necessários. Retorno: " + bodyBusca);
    }

    // 3. Montar e disparar a Emissão do Contrato
    const urlEmissao = "https://api.hinova.com.br/api/sga/v2/termo/emitir";
    
    const payloadEmissao = {
      "codigo_termo": parseInt(codigoTermoSGA),
      "codigo_associado": parseInt(veiculo.codigo_associado),
      "codigo_veiculo": parseInt(veiculo.codigo_veiculo),
      // FORÇANDO A API A APENAS GERAR O LINK DO PDF PARA OS TÉCNICOS
      "aceite_digital": "N",
      "enviar_email": "N",
      "enviar_sms": "N"
    };

    const opcoesEmissao = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + tokenSessao 
      },
      "payload": JSON.stringify(payloadEmissao),
      "muteHttpExceptions": true
    };

    const resEmissao = UrlFetchApp.fetch(urlEmissao, opcoesEmissao);
    const statusEmissao = resEmissao.getResponseCode();
    const bodyEmissao = resEmissao.getContentText();

    let jsonEmissao = JSON.parse(bodyEmissao);

    // 4. Validação do Retorno Final
    if (statusEmissao === 200 || statusEmissao === 201) {
       if (jsonEmissao.link_pdf) {
         return { sucesso: true, link: jsonEmissao.link_pdf };
       } else {
         throw new Error("SGA retornou sucesso, mas não enviou o link do PDF. Retorno: " + bodyEmissao);
       }
    } else {
       throw new Error("Erro na emissão. Status: " + statusEmissao + " | Erro SGA: " + (jsonEmissao.message || bodyEmissao));
    }

  } catch (e) {
    return { erro: e.message };
  }
}