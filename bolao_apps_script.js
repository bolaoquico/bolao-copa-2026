// ============================================================
// BOLÃO COPA 2026 – Apps Script completo
// ============================================================

const API_KEY  = 'b3ae7819475b1c36a30b5fb78e771843';
const API_HOST = 'v3.football.api-sports.io';
const LEAGUE_ID = 1;
const SEASON    = 2026;

// ============================================================
// 1. RECEBER PALPITES DO SITE (POST)
// ============================================================
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Palpites');

    // Se já existe palpite do mesmo e-mail, substitui
    const existentes = sheet.getDataRange().getValues();
    for (let i = 1; i < existentes.length; i++) {
      if (existentes[i][2] === data.email) {
        sheet.deleteRow(i + 1);
        break;
      }
    }

    const row = [new Date(), data.nome, data.email, data.telefone];
    for (let i = 1; i <= 72; i++) {
      row.push(data['j' + i] || '');
    }
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 2. BUSCAR RESULTADOS DA API-FOOTBALL
// ============================================================
function buscarResultados() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Jogos');
  const dados = sheet.getDataRange().getValues();

  let atualizou = false;

  for (let i = 1; i < dados.length; i++) {
    const status = String(dados[i][8]).trim();
    if (status === 'Finalizado') continue;

    const timeCasa = dados[i][2];
    const timeFora = dados[i][3];
    const dataStr  = dados[i][5]; // DD/MM/YYYY

    const partes   = String(dataStr).split('/');
    const dataJogo = new Date(`${partes[2]}-${partes[1]}-${partes[0]}`);
    const hoje     = new Date();
    hoje.setHours(0, 0, 0, 0);
    if (dataJogo > hoje) continue;

    const resultado = buscarJogoNaAPI(timeCasa, timeFora, dataJogo);
    if (resultado !== null) {
      sheet.getRange(i + 1, 7).setValue(resultado.casa);
      sheet.getRange(i + 1, 8).setValue(resultado.fora);
      sheet.getRange(i + 1, 9).setValue('Finalizado');
      atualizou = true;
      Logger.log(`Atualizado: ${timeCasa} ${resultado.casa} x ${resultado.fora} ${timeFora}`);
    }
  }

  if (atualizou) calcularClassificacao();

  // Atualiza timestamp na aba Config
  const sheetConfig = ss.getSheetByName('Config');
  sheetConfig.getRange('B2').setValue(new Date());
}

function buscarJogoNaAPI(timeCasa, timeFora, dataJogo) {
  try {
    const dataFormatada = Utilities.formatDate(dataJogo, 'UTC', 'yyyy-MM-dd');
    const url = `https://${API_HOST}/fixtures?date=${dataFormatada}&league=${LEAGUE_ID}&season=${SEASON}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'x-apisports-key': API_KEY },
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    const jogos = json.response || [];

    for (const jogo of jogos) {
      const homeNome   = jogo.teams.home.name;
      const awayNome   = jogo.teams.away.name;
      const statusJogo = jogo.fixture.status.short;

      if (statusJogo !== 'FT') continue;

      if (nomeParecido(homeNome, timeCasa) && nomeParecido(awayNome, timeFora)) {
        return { casa: jogo.goals.home, fora: jogo.goals.away };
      }
    }
  } catch (err) {
    Logger.log('Erro ao buscar jogo: ' + err.message);
  }
  return null;
}

function nomeParecido(nomeAPI, nomeLocal) {
  const n = s => s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z]/g, '');
  const a = n(nomeAPI);
  const b = n(nomeLocal);
  return a.includes(b.substring(0, 5)) || b.includes(a.substring(0, 5));
}

// ============================================================
// 3. CALCULAR PONTUAÇÃO
// ============================================================
function calcularPontos(palCasa, palFora, resCasa, resFora) {
  const pc = parseInt(palCasa), pf = parseInt(palFora);
  const rc = parseInt(resCasa), rf = parseInt(resFora);
  if (isNaN(pc) || isNaN(pf) || isNaN(rc) || isNaN(rf)) return 0;

  // 7 pts – escore em cheio
  if (pc === rc && pf === rf) return 7;

  const palEmpate = (pc === pf);
  const resEmpate = (rc === rf);

  // 2 pts – apostou empate, deu empate, placar errado
  if (palEmpate && resEmpate) return 2;

  if (!resEmpate) {
    const resVenc = rc > rf ? 'casa' : 'fora';
    const palVenc = pc > pf ? 'casa' : (pc < pf ? 'fora' : 'empate');

    // 5 pts – acertou escore do vencedor (mas não do vencido)
    if (resVenc === 'casa' && pc === rc && pf !== rf) return 5;
    if (resVenc === 'fora' && pf === rf && pc !== rc) return 5;

    // 4 pts – acertou escore do vencido (mas não do vencedor)
    if (resVenc === 'casa' && pf === rf && pc !== rc) return 4;
    if (resVenc === 'fora' && pc === rc && pf !== rf) return 4;

    // 1 pt – escore invertido
    if (pc === rf && pf === rc) return 1;

    // 3 pts – acertou apenas o vencedor
    if (palVenc === resVenc) return 3;
  }

  return 0;
}

// ============================================================
// 4. CALCULAR CLASSIFICAÇÃO
// ============================================================
function calcularClassificacao() {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const sheetJogos    = ss.getSheetByName('Jogos');
  const sheetPalpites = ss.getSheetByName('Palpites');
  const sheetClass    = ss.getSheetByName('Classificacao');

  const jogos    = sheetJogos.getDataRange().getValues();
  const palpites = sheetPalpites.getDataRange().getValues();

  // Mapa de resultados finalizados
  const resultados = {};
  let finalizados  = 0;
  for (let i = 1; i < jogos.length; i++) {
    const id     = parseInt(jogos[i][0]);
    const casa   = jogos[i][6];
    const fora   = jogos[i][7];
    const status = String(jogos[i][8]).trim();
    if (status === 'Finalizado' && casa !== '' && fora !== '') {
      resultados[id] = { casa, fora };
      finalizados++;
    }
  }

  // Pontuação de cada apostador
  const classificacao = [];
  for (let i = 1; i < palpites.length; i++) {
    const linha = palpites[i];
    if (!linha[1]) continue;

    let totalPontos = 0;
    let acertos     = 0;

    for (let j = 1; j <= 72; j++) {
      const res = resultados[j];
      if (!res) continue;

      // J1 está no índice 4 (0=Timestamp,1=Nome,2=Email,3=Tel,4=J1...)
      const palpite = String(linha[3 + j] || '').trim();
      if (!palpite || !palpite.includes('-')) continue;

      const [palCasa, palFora] = palpite.split('-').map(v => v.trim());
      const pts = calcularPontos(palCasa, palFora, res.casa, res.fora);
      totalPontos += pts;
      if (pts === 7) acertos++;
    }

    classificacao.push({ nome: linha[1], email: linha[2], pontos: totalPontos, acertos });
  }

  // Ordena por pontos, desempata por acertos exatos
  classificacao.sort((a, b) => b.pontos - a.pontos || b.acertos - a.acertos);

  // Grava classificação
  sheetClass.clearContents();
  sheetClass.appendRow(['Posição', 'Nome', 'E-mail', 'Pontos', 'Placares Exatos', 'Jogos Apurados']);
  classificacao.forEach((ap, idx) => {
    sheetClass.appendRow([idx + 1, ap.nome, ap.email, ap.pontos, ap.acertos, finalizados]);
  });

  // Formata cabeçalho
  sheetClass.getRange(1, 1, 1, 6)
    .setFontWeight('bold')
    .setBackground('#006633')
    .setFontColor('#ffffff');

  Logger.log(`Classificação: ${classificacao.length} apostadores, ${finalizados} jogos apurados.`);
}

// ============================================================
// 5. PUBLICAR JSON PARA O SITE DE CLASSIFICAÇÃO (GET)
// ============================================================
function doGet(e) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheetClass = ss.getSheetByName('Classificacao');
  const sheetJogos = ss.getSheetByName('Jogos');

  const dados = sheetClass.getDataRange().getValues();
  const classificacao = [];
  for (let i = 1; i < dados.length; i++) {
    if (!dados[i][1]) continue;
    classificacao.push({
      posicao: dados[i][0],
      nome:    dados[i][1],
      pontos:  dados[i][3],
      acertos: dados[i][4]
    });
  }

  const jogos = sheetJogos.getDataRange().getValues();
  let finalizados = 0;
  for (let i = 1; i < jogos.length; i++) {
    if (String(jogos[i][8]).trim() === 'Finalizado') finalizados++;
  }

  const output = {
    atualizadoEm:  new Date().toISOString(),
    jogosApurados: finalizados,
    totalJogos:    72,
    classificacao
  };

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 6. CONFIGURAR GATILHO AUTOMÁTICO (rode uma única vez)
// ============================================================
function configurarGatilho() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('buscarResultados')
    .timeBased()
    .everyHours(2)
    .create();

  Logger.log('Gatilho configurado: buscarResultados a cada 2 horas.');
}
