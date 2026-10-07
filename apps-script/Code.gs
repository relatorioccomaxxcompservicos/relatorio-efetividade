/**
 * Importador SAR2G -> aba EFETIVIDADE  (versão 3)
 * ------------------------------------------------
 * O arquivo do SAR2G é lido e agregado no NAVEGADOR (ver Dialog.html), o que evita o limite de
 * 6 minutos do Apps Script. Este script só recebe as linhas prontas e as grava na planilha.
 *
 * Proteções:
 *  - grava SÓ os dias presentes no arquivo (as demais linhas da aba não são lidas nem regravadas), o que
 *    mantém a importação rápida mesmo com a aba preenchida até o fim do ano;
 *  - guarda as linhas substituídas na aba oculta "EFETIVIDADE_backup" antes de alterar;
 *  - escreve as linhas novas ANTES de apagar as antigas (se algo falhar, nada se perde);
 *  - confere o que foi gravado e avisa se algo não bateu.
 *
 * Linhas-modelo: o menu "Preparar datas futuras" cria, até a data escolhida, uma linha por posto+cargo e dia
 * com DATA, DIA, EMPRESA, CÓDIGOS, POSTO e CARGO preenchidos e o resto (H em diante) em branco. Ao importar um
 * dia, as linhas-modelo dele são substituídas pelos dados reais.
 *
 * Instalação: Extensões > Apps Script > cole este arquivo (script.gs), o Dialog.html e o appsscript.json.
 */

const ABA_DADOS = 'EFETIVIDADE';
const ABA_BACKUP = 'EFETIVIDADE_backup';
const ABA_CONFIG = 'SAR2G_config';     // oculta: guarda a última linha com dados reais (lida pelo dashboard)
const LIMITE_CELULAS = 10000000;       // limite de células de uma planilha Google
const LINHA_DADOS_INICIO = 4;          // linhas 1-3 são título/cabeçalho
const TOTAL_COLS = 31;                 // A..U (originais) + V..AE (novas, vindas do SAR2G)
const COL_NOVAS_INICIO = 22;           // coluna V
const CABECALHOS_NOVAS = ['COD_CLIENTE', 'COD_LOCAL', 'COD_CARGO', 'COD_GESTOR', 'COD_AREA_SUPERVISAO',
                          'AREA_SUPERVISAO', 'COD_ESCALA', 'DESC_ESCALA', 'SIGLA_ESCALA', 'TPCLIENTE'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SAR2G')
    .addItem('Importar relatório (.xlsx)', 'abrirDialogoImportacao')
    .addItem('Preparar datas futuras (linhas em branco)', 'prepararDatasFuturas')
    .addToUi();
}

function abrirDialogoImportacao() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog').setWidth(600).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Importar relatório do SAR2G');
}

/** Ponto de entrada chamado pelo Dialog.html com as linhas já agregadas (JSON). */
function gravarLinhas(jsonLinhas, nomeArquivo, diasInteiros) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Já existe uma importação em andamento. Aguarde terminar.');
  try {
    return gravar_(jsonLinhas, nomeArquivo, diasInteiros !== false);
  } finally {
    lock.releaseLock();
  }
}

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function chaveData_(v, tz) {
  if (v instanceof Date) {
    // Datas gravadas por versões anteriores do importador ficaram deslocadas: a meia-noite de São Paulo
    // aparecia como "dia anterior, 20:00" no fuso da planilha. Arredondar para a meia-noite mais
    // próxima (no fuso da planilha) devolve o dia correto tanto para essas quanto para datas digitadas.
    const h = +Utilities.formatDate(v, tz, 'H');
    const d = h >= 12 ? new Date(v.getTime() + 12 * 3600 * 1000) : v;
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : s;
}

/** 'AAAA-MM-DD' -> data à meia-noite NO FUSO DA PLANILHA (e não no fuso do script). */
function dataNaPlanilha_(iso, tz) {
  return Utilities.parseDate(iso, tz, 'yyyy-MM-dd');
}

const OBS_AUTOMATICAS = ['Cobertura atípica entre funções', 'Vaga sem titular fixo (coberta por reserva/FT)'];

/** Devolve só o que foi digitado à mão numa observação (remove as frases geradas pelo importador). */
function obsManual_(txt) {
  return String(txt == null ? '' : txt).split(';').map(function (x) { return x.trim(); })
    .filter(function (x) { return x && OBS_AUTOMATICAS.indexOf(x) < 0; }).join('; ');
}

function chaveLinha_(r, tz) {
  // sem empresa de propósito (ver comentário no passo 5 de gravar_): um posto pode trocar de empresa
  // entre importações, e a anotação manual precisa continuar sendo encontrada mesmo assim.
  return chaveData_(r[0], tz) + '|' + norm_(r[5]) + '|' + norm_(r[6]);
}

function limpaErro_(v) {
  return (typeof v === 'string' && /^#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|ERROR!)$/.test(v)) ? '' : v;
}

const DIAS_PT = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];

function diaSemanaPt_(iso) {
  const p = iso.split('-');
  return DIAS_PT[new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay()];
}

function somaDias_(iso, n) {
  const p = iso.split('-'), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.toISOString().slice(0, 10);
}

function totalCelulas_(ss) {
  return ss.getSheets().reduce(function (s, sh) { return s + sh.getMaxRows() * sh.getMaxColumns(); }, 0);
}

function garanteColunas_(aba) {
  if (aba.getMaxColumns() < TOTAL_COLS) aba.insertColumnsAfter(aba.getMaxColumns(), TOTAL_COLS - aba.getMaxColumns());
  aba.getRange(3, COL_NOVAS_INICIO, 1, CABECALHOS_NOVAS.length).setValues([CABECALHOS_NOVAS]);
  if (CABECALHOS_NOVAS.length > 2) {
    aba.getRange(3, COL_NOVAS_INICIO + 1, 1, 1)
      .copyTo(aba.getRange(3, COL_NOVAS_INICIO + 2, 1, CABECALHOS_NOVAS.length - 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
}

/** Datas (AAAA-MM-DD) da coluna A, a partir da linha 4. Linhas sem data viram ''. */
function lerDatas_(aba, tz) {
  const n = Math.max(aba.getLastRow() - LINHA_DADOS_INICIO + 1, 0);
  if (!n) return [];
  return aba.getRange(LINHA_DADOS_INICIO, 1, n, 1).getValues()
    .map(function (r) { return r[0] === '' || r[0] == null ? '' : chaveData_(r[0], tz); });
}

/** Linha (na planilha) da última linha com dados reais: coluna H (quadro diurno) preenchida. */
function atualizarUltimaLinha_(ss, aba) {
  const n = Math.max(aba.getLastRow() - LINHA_DADOS_INICIO + 1, 0);
  let ult = LINHA_DADOS_INICIO - 1;
  if (n) {
    const h = aba.getRange(LINHA_DADOS_INICIO, 8, n, 1).getValues();
    for (let i = n - 1; i >= 0; i--) if (h[i][0] !== '' && h[i][0] != null) { ult = LINHA_DADOS_INICIO + i; break; }
  }
  let cfg = ss.getSheetByName(ABA_CONFIG);
  if (!cfg) { cfg = ss.insertSheet(ABA_CONFIG); cfg.hideSheet(); ss.setActiveSheet(aba); }
  cfg.getRange(1, 1, 1, 2).setValues([['ultima_linha_com_dados', ult]]);
  return ult;
}

/** Insere `qtd` linhas a partir da linha `linha` (empurrando o que estiver ali para baixo). */
function inserirLinhas_(aba, linha, qtd) {
  if (linha > aba.getMaxRows()) aba.insertRowsAfter(aba.getMaxRows(), qtd + (linha - aba.getMaxRows() - 1));
  else aba.insertRowsBefore(linha, qtd);
}

/** Copia a formatação de uma linha de dados existente (fora do trecho) para o trecho [linha, linha+qtd). */
function formatar_(aba, linha, qtd) {
  const ultima = aba.getLastRow();
  let ref = linha + qtd <= ultima ? linha + qtd : (linha - 1 >= LINHA_DADOS_INICIO ? linha - 1 : 0);
  if (!ref) return;
  aba.getRange(ref, 1, 1, TOTAL_COLS)
    .copyTo(aba.getRange(linha, 1, qtd, TOTAL_COLS), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

/** Apaga as linhas (índices relativos à linha 4, já em ordem crescente) de baixo para cima, por trechos. */
function apagarLinhas_(aba, indices, deslocamento) {
  for (let i = indices.length - 1; i >= 0;) {
    let j = i;
    while (j > 0 && indices[j - 1] === indices[j] - 1) j--;
    aba.deleteRows(LINHA_DADOS_INICIO + indices[j] + deslocamento, indices[i] - indices[j] + 1);
    i = j - 1;
  }
}

function gravar_(jsonLinhas, nomeArquivo, diasInteiros) {
  const novas = JSON.parse(jsonLinhas);
  if (!Array.isArray(novas) || novas.length === 0) throw new Error('Nenhuma linha recebida para gravar.');
  novas.forEach(function (r, i) {
    if (!Array.isArray(r) || r.length !== TOTAL_COLS) {
      throw new Error('Linha ' + (i + 1) + ' com formato inesperado (' + (r && r.length) + ' colunas; esperado ' + TOTAL_COLS + ').');
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_DADOS);
  if (!aba) throw new Error('Aba "' + ABA_DADOS + '" não encontrada.');
  const tz = ss.getSpreadsheetTimeZone();
  garanteColunas_(aba);

  const porData = {}, chavesImport = {};
  novas.forEach(function (r) {
    (porData[r[0]] = porData[r[0]] || []).push(r);
    chavesImport[r[0] + '|' + norm_(r[5]) + '|' + norm_(r[6])] = true;
  });
  const datas = Object.keys(porData).sort();
  // substituir esta linha antiga? dias inteiros (padrão) = todas as linhas do dia; senão só posto+cargo do arquivo.
  // A empresa fica fora da chave de propósito: um posto pode trocar de empresa entre importações.
  const substituir = function (d, r) { return diasInteiros || !!chavesImport[d + '|' + norm_(r[5]) + '|' + norm_(r[6])]; };

  // 1) cópia de segurança: só as linhas que serão substituídas (leve, mesmo com a aba cheia de linhas-modelo)
  const colA = lerDatas_(aba, tz), alvo = {};
  datas.forEach(function (d) { alvo[d] = true; });
  const idxAlvo = [];
  colA.forEach(function (d, i) { if (alvo[d]) idxAlvo.push(i); });
  const backup = [];
  if (idxAlvo.length) {
    const ini = idxAlvo[0], fim = idxAlvo[idxAlvo.length - 1];
    const bloco = aba.getRange(LINHA_DADOS_INICIO + ini, 1, fim - ini + 1, TOTAL_COLS).getValues();
    idxAlvo.forEach(function (i) { const r = bloco[i - ini]; if (substituir(colA[i], r)) backup.push(r); });
  }
  const antigo = ss.getSheetByName(ABA_BACKUP);
  if (antigo) ss.deleteSheet(antigo);
  const bk = ss.insertSheet(ABA_BACKUP);
  bk.getRange(1, 1, 1, 2).setValues([['Linhas substituídas na importação de ' + (nomeArquivo || ''), new Date()]]);
  if (backup.length) {
    if (bk.getMaxColumns() < TOTAL_COLS) bk.insertColumnsAfter(bk.getMaxColumns(), TOTAL_COLS - bk.getMaxColumns());
    if (bk.getMaxRows() < backup.length + 1) bk.insertRowsAfter(bk.getMaxRows(), backup.length + 1 - bk.getMaxRows());
    bk.getRange(2, 1, backup.length, TOTAL_COLS).setValues(backup);
  }
  bk.hideSheet();
  ss.setActiveSheet(aba);

  // 2) cada dia, do último para o primeiro (assim as alterações não deslocam os dias ainda não processados)
  let substituidas = 0, mantidasNoDia = 0, anotacoesPreservadas = 0;
  const esperado = {};
  datas.slice().reverse().forEach(function (d) {
    const col = lerDatas_(aba, tz), idx = [];
    let depois = -1;
    col.forEach(function (x, i) { if (x === d) idx.push(i); else if (depois < 0 && x && x > d) depois = i; });

    let antigas = [];
    if (idx.length) {
      const ini = idx[0], fim = idx[idx.length - 1];
      aba.getRange(LINHA_DADOS_INICIO + ini, 1, fim - ini + 1, TOTAL_COLS).breakApart();
      const bloco = aba.getRange(LINHA_DADOS_INICIO + ini, 1, fim - ini + 1, TOTAL_COLS).getValues();
      antigas = idx.map(function (i) { return bloco[i - ini]; });
    }
    const mantidas = [], manuais = {};
    antigas.forEach(function (r) {
      if (!substituir(d, r)) { const c = r.map(limpaErro_); c[0] = dataNaPlanilha_(d, tz); mantidas.push(c); return; }
      substituidas++;
      const m = { s: r[18], t: obsManual_(r[19]), u: r[20] };                    // reserva/observação digitadas à mão
      if (m.s !== '' || m.t !== '' || m.u !== '') manuais[norm_(r[5]) + '|' + norm_(r[6])] = m;
    });
    mantidasNoDia += mantidas.length;
    const novasDia = porData[d].map(function (r) {
      const c = r.slice();
      c[0] = dataNaPlanilha_(d, tz);
      const m = manuais[norm_(c[5]) + '|' + norm_(c[6])];
      if (m) {                                                                     // reimportação: não perde o que a equipe escreveu
        if (c[18] === '' && m.s !== '') c[18] = m.s;
        if (c[20] === '' && m.u !== '') c[20] = m.u;
        if (m.t) c[19] = [m.t, c[19]].filter(function (x) { return x; }).join('; ');
        anotacoesPreservadas++;
      }
      return c;
    });
    const finais = mantidas.concat(novasDia).map(function (r) { return { r: r, k: [norm_(r[2]), norm_(r[5]), norm_(r[6])] }; });
    finais.sort(function (a, b) { for (let i = 0; i < 3; i++) if (a.k[i] !== b.k[i]) return a.k[i] < b.k[i] ? -1 : 1; return 0; });
    const linhas = finais.map(function (x) { return x.r; });
    esperado[d] = linhas.length;

    // escreve as linhas novas ANTES de apagar as antigas: se algo falhar, nada se perde
    const pos = idx.length ? idx[0] : (depois >= 0 ? depois : col.length);
    const linha = LINHA_DADOS_INICIO + pos;
    inserirLinhas_(aba, linha, linhas.length);
    try {
      aba.getRange(linha, 1, linhas.length, TOTAL_COLS).setValues(linhas);
    } catch (e) {
      aba.deleteRows(linha, linhas.length);                                        // desfaz as linhas em branco inseridas
      throw e;
    }
    formatar_(aba, linha, linhas.length);
    apagarLinhas_(aba, idx, linhas.length);
  });

  const ultimaLinha = atualizarUltimaLinha_(ss, aba);
  SpreadsheetApp.flush();

  // 3) confere: cada dia importado tem exatamente as linhas esperadas
  const contagem = {};
  lerDatas_(aba, tz).forEach(function (d) { if (esperado[d] != null) contagem[d] = (contagem[d] || 0) + 1; });
  const errados = datas.filter(function (d) { return contagem[d] !== esperado[d]; });
  if (errados.length) {
    throw new Error('A conferência final não bateu nos dias ' + errados.join(', ') + '. As linhas anteriores estão na aba oculta "' + ABA_BACKUP + '" (clique com o botão direito nas abas > Mostrar planilhas ocultas).');
  }

  const empresas = {};
  novas.forEach(function (r) { empresas[r[2]] = 1; });
  return 'Importação concluída' + (nomeArquivo ? ' (' + nomeArquivo + ')' : '') + '.\n' +
    novas.length + ' linhas gravadas de ' + Object.keys(empresas).length + ' empresas, ' +
    datas.length + ' dias (' + datas[0] + ' a ' + datas[datas.length - 1] + ').\n' +
    substituidas + ' linhas antigas (ou em branco) substituídas' + (mantidasNoDia ? '; ' + mantidasNoDia + ' de outras empresas mantidas nesses dias' : '') + '.\n' +
    (anotacoesPreservadas ? anotacoesPreservadas + ' linhas mantiveram as anotações digitadas à mão.\n' : '') +
    'Demais dias da aba não foram alterados. Última linha com dados: ' + ultimaLinha + '.\n' +
    'Linhas substituídas guardadas na aba oculta "' + ABA_BACKUP + '".';
}

/* ------------------------------------------------------------------------------------------------------
 * Linhas-modelo para datas futuras
 * ---------------------------------------------------------------------------------------------------- */
function prepararDatasFuturas() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Preparar datas futuras',
    'Criar linhas em branco (data, posto e cargo) até qual data? Use dd/mm/aaaa.', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(resp.getResponseText().trim());
  if (!m) { ui.alert('Data inválida. Use o formato dd/mm/aaaa, por exemplo 31/12/2026.'); return; }
  ui.alert(criarModelo_(m[3] + '-' + m[2] + '-' + m[1]));
}

/**
 * Cria, do dia seguinte à última data da aba até `ateIso`, uma linha por empresa+posto+cargo e dia, usando os
 * postos e cargos dos últimos 7 dias com dados reais. Preenche só DATA, DIA, EMPRESA, COD. EMP, COD. POSTO,
 * POSTO e CARGO; o resto fica em branco até a importação. Se o tempo acabar, para e pode ser executado de novo
 * (continua de onde parou).
 */
function criarModelo_(ateIso) {
  const inicio = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_DADOS);
  if (!aba) throw new Error('Aba "' + ABA_DADOS + '" não encontrada.');
  const tz = ss.getSpreadsheetTimeZone();
  garanteColunas_(aba);
  atualizarUltimaLinha_(ss, aba);          // o dashboard passa a ler só até aqui, ignorando as linhas-modelo

  const n = Math.max(aba.getLastRow() - LINHA_DADOS_INICIO + 1, 0);
  if (!n) return 'A aba está vazia. Importe pelo menos um dia antes de preparar as datas futuras.';
  const vals = aba.getRange(LINHA_DADOS_INICIO, 1, n, 8).getValues();   // A..H
  let ultimaData = '';
  const comDados = {};
  vals.forEach(function (r) {
    if (r[0] === '' || r[0] == null) return;
    const d = chaveData_(r[0], tz);
    if (d > ultimaData) ultimaData = d;
    if (r[7] !== '' && r[7] != null) comDados[d] = true;
  });
  const diasReais = Object.keys(comDados).sort().slice(-7);
  if (!diasReais.length) return 'Não há dias com dados importados para servir de padrão.';
  if (ultimaData >= ateIso) return 'A aba já vai até ' + ultimaData + '. Nada a criar.';

  // postos e cargos dos últimos 7 dias importados
  const usar = {}, combos = {};
  diasReais.forEach(function (d) { usar[d] = true; });
  vals.forEach(function (r) {
    if (r[0] === '' || r[0] == null || !usar[chaveData_(r[0], tz)] || !r[5] || !r[6]) return;
    const k = norm_(r[2]) + '|' + norm_(r[5]) + '|' + norm_(r[6]);
    combos[k] = { emp: r[2], codEmp: r[3], codPosto: r[4], posto: r[5], cargo: r[6], k: k };
  });
  const lista = Object.keys(combos).sort().map(function (k) { return combos[k]; });

  const dias = [];
  for (let d = somaDias_(ultimaData, 1); d <= ateIso; d = somaDias_(d, 1)) dias.push(d);
  const novasLinhas = dias.length * lista.length;
  const projetado = totalCelulas_(ss) + Math.max(0, aba.getLastRow() + novasLinhas - aba.getMaxRows()) * aba.getMaxColumns();
  if (projetado > LIMITE_CELULAS * 0.95) {
    return 'Não criei as linhas: a planilha passaria de ' + Math.round(projetado / 1e6 * 10) / 10 +
      ' milhões de células (limite do Google: 10 milhões). Escolha uma data final mais próxima.';
  }

  let criadosAte = '', linhasCriadas = 0;
  const LOTE = 15000;
  for (let i = 0; i < dias.length;) {
    const lote = [];
    let j = i;
    while (j < dias.length && (lote.length === 0 || lote.length + lista.length <= LOTE)) {
      const data = dataNaPlanilha_(dias[j], tz), dia = diaSemanaPt_(dias[j]);
      lista.forEach(function (c) { lote.push([data, dia, c.emp, c.codEmp, c.codPosto, c.posto, c.cargo]); });
      j++;
    }
    const linha = aba.getLastRow() + 1;
    if (linha + lote.length - 1 > aba.getMaxRows()) aba.insertRowsAfter(aba.getMaxRows(), linha + lote.length - 1 - aba.getMaxRows());
    aba.getRange(linha, 1, lote.length, 7).setValues(lote);
    formatar_(aba, linha, lote.length);
    SpreadsheetApp.flush();
    linhasCriadas += lote.length; criadosAte = dias[j - 1]; i = j;
    if (Date.now() - inicio > 4.5 * 60 * 1000 && i < dias.length) {
      return 'Criadas ' + linhasCriadas + ' linhas, até ' + criadosAte + '. O tempo do Google acabou antes do fim: ' +
        'rode "Preparar datas futuras" de novo com a mesma data para continuar de onde parou.';
    }
  }
  return 'Pronto: ' + linhasCriadas + ' linhas em branco criadas, de ' + dias[0] + ' a ' + criadosAte + ' (' +
    lista.length + ' postos/cargos por dia, com base nos últimos ' + diasReais.length + ' dias importados).\n' +
    'Ao importar um dia, as linhas dele são substituídas pelos dados reais. Postos novos que aparecerem no ' +
    'relatório entram automaticamente; postos que deixarem de existir saem do dia quando ele for importado.';
}
