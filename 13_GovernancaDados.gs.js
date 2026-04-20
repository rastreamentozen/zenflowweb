// ====================================================================================
// 🧠 ARQUIVO: 13_GovernancaDados.gs (Gestão de Histórico e Virada de Mês)
// ====================================================================================

function executarViradaDeMes() {
  const ssVigente = SpreadsheetApp.openById(ID_PLANILHA_VIGENTE);
  const ssHistorico = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
  
  const dataAnterior = new Date();
  dataAnterior.setMonth(dataAnterior.getMonth() - 1);
  
  const nomesMeses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  const nomeMesAnterior = `${nomesMeses[dataAnterior.getMonth()]} ${dataAnterior.getFullYear()}`;
  const nomeMesNovo = `${nomesMeses[new Date().getMonth()]} ${new Date().getFullYear()}`;

  let abaParaArquivar = ssVigente.getSheets()[0];
  
  // 1. Arquivamento
  abaParaArquivar.copyTo(ssHistorico).setName(nomeMesAnterior);
  
  // 2. Limpeza da Planilha Vigente 
  const ultimaLinha = abaParaArquivar.getLastRow();
  const ultimaColuna = abaParaArquivar.getLastColumn();
  
  if (ultimaLinha > 1) {
    abaParaArquivar.getRange(2, 1, ultimaLinha - 1, ultimaColuna).clearContent();
  }
  
  // 3. Batismo do novo mês
  abaParaArquivar.setName(nomeMesNovo);
}