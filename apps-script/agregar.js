/**
 * Agregação do relatório de ponto do SAR2G -> linhas da aba EFETIVIDADE.
 * Função pura (sem APIs do Google): roda no navegador (dentro do Dialog.html) e no Node (testes).
 *
 * "Posto" = LOCAL DE SERVIÇO (NOMELOCAL), não o cliente (NOMECLIENTE): um cliente pode ter vários
 * locais físicos diferentes (ex.: um cliente com sede + filiais).
 *
 * Regras (validadas com a operação):
 *  - QE (quadro esperado no dia), por empresa+local+cargo+turno = registros tipo EFETIVO com situação em
 *    {TRABALHO, FALTA, LIB. PAR. FUNC, LIB. PAR. COB}
 *  - Presente = registros (qualquer tipo) com situação = TRABALHO
 *  - Reserva técnica acionada = registros tipo COBERTURA ou FT com situação = TRABALHO
 *  - Vaga sem titular (EFETIVO nunca aparece, mas há gente trabalhando): o QE acompanha os presentes
 *  - Falta não coberta = QE - presentes; se outro cargo do MESMO empresa+local+turno tiver sobra de
 *    presença no dia, essa sobra fecha o déficit antes (cobertura atípica entre funções)
 *  - Turno pela hora de entrada (HRENTRADA): diurno 05:00-17:59, noturno caso contrário
 *
 *  Fora do escopo de efetividade operacional (excluídos da planilha):
 *  - cliente == RESERVA TÉCNICA (qualquer local: bolsas regionais, abandono, afastados, processo
 *    trabalhista, fracionados etc.) — a cobertura que essas pessoas fazem em outros postos continua
 *    contando normalmente lá, só a "casa" da reserva é que não vira linha na planilha.
 *  - cliente == MAXXSEG (administrativo: RH, DP, financeiro, supervisão, suprimentos...)
 *  - cliente == GOTHAM com local em {PORTARIA REMOTA, ADM} (administrativo da Gotham)
 *  - empresa+local+cargo sem NENHUM quadro e NENHUMA presença em TODO o período do arquivo
 *
 * Entrada: array de objetos {NOME_DA_COLUNA: valor} (como sheet_to_json do SheetJS, com raw:true).
 * Saída:   { linhas: [[...31 colunas...]], resumo: {...} }
 */
(function (root) {
  var QE_STATUS = { 'TRABALHO': 1, 'FALTA': 1, 'LIB. PAR. FUNC': 1, 'LIB. PAR. COB': 1 };
  var RESERVA_TIPOS = { 'COBERTURA': 1, 'FT': 1, 'DOBRA': 1, 'CONVOCAÇÃO': 1 };   // todos ocupam uma vaga do quadro
  var EXTRAS = ['COD_CARGO', 'COD_GESTOR', 'CODAREASUPERVISAO', 'AREASUPERVISAO',
                'COD_ESCALA', 'DESCESCALA', 'SIGLA', 'TPCLIENTE'];
  var DIAS = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
  var EPOCH = Date.UTC(1899, 11, 30);
  var GOTHAM_ADM = { 'PORTARIA REMOTA': 1, 'ADM': 1 };

  /**
   * Lê o relatório de MOVIMENTAÇÃO (opcional, mesmo período do ponto) e extrai três coisas:
   *  - feriasPorRE: RE do funcionário -> lista de {ini,fim} das férias (do sistema, não do ponto,
   *    porque no ponto um dia de férias vem com a situação em branco, indistinguível de outros motivos)
   *  - empresaPorLocal: NOMELOCAL -> empresa DONA do contrato (a "vaga"), não a empresa de quem
   *    cobriu naquele dia (que pode ser de outra empresa, ex.: reserva emprestada)
   *  - descobertoPorChave: 'empresa|local|data|turno' -> quantidade de eventos "POSTO DESCOBERTO"
   *    confirmados pelo próprio sistema, usados só para conferência (não altera o cálculo sozinho)
   */
  function parseMovimentacao(rows) {
    var feriasPorRE = {}, contagemLocalEmpresa = {}, descobertoPorChave = {}, descobertos = [], vistoDesc = {};
    rows.forEach(function (r) {
      var local = str(r.NOMELOCAL), emp = str(r.NOMEEMPRESA);
      if (local && emp) {
        var c = contagemLocalEmpresa[local] || (contagemLocalEmpresa[local] = {});
        c[emp] = (c[emp] || 0) + 1;
      }
      if (str(r.TPMOVIMENTO) === 'FÉRIAS' && !vazio(r.RE)) {
        var re = String(r.RE), ini = dataStr(r.DTINICIO), fim = vazio(r.DTFIM) ? ini : dataStr(r.DTFIM);
        (feriasPorRE[re] = feriasPorRE[re] || []).push({ ini: ini, fim: fim });
      }
      if (str(r.TPMOVIMENTO) === 'POSTO DESCOBERTO') {
        var tn = str(r.DESCTURNO) === 'NOTURNO' ? 'N' : 'D';
        var k = [emp, local, dataStr(r.DTINICIO), tn].join('|');
        descobertoPorChave[k] = (descobertoPorChave[k] || 0) + 1;
        var vaga = vazio(r.VAGAITEMCONTRATO) ? '' : String(r.VAGAITEMCONTRATO);
        var kd = vaga ? dataStr(r.DTINICIO) + '|' + vaga : k + '|' + str(r.DESCVAGA) + '|' + descobertos.length;
        if (!vistoDesc[kd]) {
          vistoDesc[kd] = 1;
          descobertos.push({ data: dataStr(r.DTINICIO), cliente: str(r.NOMECLIENTE), local: local, emp: emp,
                             cargo: str(r.DESCVAGA), turno: tn, vaga: vaga });
        }
      }
    });
    var empresaPorLocal = {};
    Object.keys(contagemLocalEmpresa).forEach(function (local) {
      var c = contagemLocalEmpresa[local], melhor = null, max = -1;
      Object.keys(c).forEach(function (e) { if (c[e] > max) { max = c[e]; melhor = e; } });
      empresaPorLocal[local] = melhor;
    });
    return { feriasPorRE: feriasPorRE, empresaPorLocal: empresaPorLocal, descobertoPorChave: descobertoPorChave, descobertos: descobertos };
  }

  function emFerias(feriasPorRE, re, data) {
    var lista = feriasPorRE[String(re)];
    if (!lista) return false;
    for (var i = 0; i < lista.length; i++) if (lista[i].ini <= data && data <= lista[i].fim) return true;
    return false;
  }

  // Feriados nacionais (as escalas "... FOLGA FER" folgam neles). Feriados estaduais/municipais não estão aqui.
  var FERIADOS = {};
  ['2026-01-01','2026-04-03','2026-04-21','2026-05-01','2026-09-07','2026-10-12','2026-11-02','2026-11-15','2026-11-20',
   '2026-12-25','2027-01-01','2027-03-26','2027-04-21','2027-05-01','2027-09-07','2027-10-12','2027-11-02','2027-11-15',
   '2027-11-20','2027-12-25'].forEach(function (d) { FERIADOS[d] = 1; });
  function diaNum(d) { var p = d.split('-'); return Math.round(Date.UTC(+p[0], +p[1] - 1, +p[2]) / 864e5); }
  function diaSemana(d) { var p = d.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); }

  /** O titular trabalharia nesse dia pela escala? true / false / null (não dá para saber). */
  function trabalhaNaEscala(escala, d, paridade) {
    var e = str(escala).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    var w = diaSemana(d), fer = !!FERIADOS[d];
    if (/12X36|DIA SIM/.test(e)) return paridade == null ? null : (diaNum(d) % 2 === paridade);   // dia sim, dia não
    if (/TROCA|INVERTE/.test(e)) {                // seg a sex fixos; fim de semana revezado (sem como saber a vez)
      if (fer) return false;
      if (w >= 1 && w <= 5) return true;
      return w === 6 && /INVERTE DF/.test(e);     // INVERTE DF: trabalha todos os sábados e reveza domingos
    }
    if (/FOLGA SAB\/DOM/.test(e)) return w >= 1 && w <= 5 && !fer;                 // 5x2
    if (/FOLGA DOM/.test(e)) return w !== 0 && !(fer && /FER/.test(e));             // 6x1 folga dom (/fer)
    return null;                                  // escalas em ciclo (4x2, 5x1, 6x2...): sem como saber
  }

  function excluido(cliente, local) {
    if (cliente === 'RESERVA TÉCNICA') return true;
    if (cliente === 'MAXXSEG') return true;
    if (cliente === 'GOTHAM' && GOTHAM_ADM[local]) return true;
    return false;
  }

  function str(v) { return v == null ? '' : String(v).trim(); }
  // situação do dia sem o ponto final (o SAR2G grava "LIB. PAR. FUNC.", "LIB. PAR. COB.", "LIB. PAR. CLI.")
  function situacao(v) { return str(v).replace(/\.$/, ''); }
  function vazio(v) { return v === '' || v == null; }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  // data do Excel (número serial) ou texto -> 'AAAA-MM-DD'
  function dataStr(v) {
    if (typeof v === 'number') {
      var d = new Date(EPOCH + Math.floor(v) * 86400000);
      return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    }
    if (v instanceof Date) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
    var s = str(v), m;
    if ((m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s))) return m[1] + '-' + m[2] + '-' + m[3];
    if ((m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s))) return m[3] + '-' + m[2] + '-' + m[1];
    return '';
  }

  // hora de entrada -> 'D' (05:00-17:59) ou 'N'; sem horário válido, assume diurno
  function turno(v) {
    var h = null, m;
    if (typeof v === 'number') h = Math.min(23, Math.floor((v - Math.floor(v)) * 24 + 1e-6));
    else if (v instanceof Date) h = v.getHours();
    else if ((m = /(\d{1,2}):(\d{2})/.exec(str(v)))) h = parseInt(m[1], 10);
    if (h == null) return 'D';
    return (h >= 5 && h < 18) ? 'D' : 'N';
  }

  function agregar(rows, movRows) {
    var grupos = {}, ordem = [], meta = {}, empresas = {};
    var mov = movRows && movRows.length ? parseMovimentacao(movRows) : null;
    var feriasContadas = 0, vagasAbertasCobertas = 0;

    // Regra por VAGA (COD_VAGA). Uma vaga entra no quadro esperado do dia quando:
    //  - o titular está escalado (trabalho, falta, saída sem autorização, liberado p/ cobertura ou férias) — já
    //    contado pelo próprio titular; ou
    //  - alguém a cobre (cobertura, folga trabalhada, dobra, convocação): se a vaga foi ocupada, ela era
    //    necessária naquele dia, mesmo sem titular (vaga a contratar / escala sem titular fixo, como o vigilante
    //    de sáb/dom/fer) ou com o titular liberado/de folga. Conta uma vez por vaga e dia.
    // férias: só contam em dia de trabalho do titular — pela escala, ou quando a vaga foi coberta ou registrada
    // como "posto descoberto" naquele dia. Para 12x36 / dia sim dia não, a vez é descoberta pelos outros dias da
    // própria pessoa ou da vaga no período.
    var cobertaNoDia = {}, descobertaNoDia = {}, parRE = {}, parVaga = {};
    rows.forEach(function (r) {
      var tipo = str(r.DESCTPCOBERTURA), d = dataStr(r.DATA), sit = situacao(r.DESCSITUACAOHOJE);
      if (!d) return;
      if (RESERVA_TIPOS[tipo] && sit === 'TRABALHO' && !vazio(r.COD_VAGA)) {
        cobertaNoDia[d + '|' + r.COD_VAGA] = 1;
        if (parVaga[String(r.COD_VAGA)] == null) parVaga[String(r.COD_VAGA)] = diaNum(d) % 2;
      }
      if (tipo === 'EFETIVO' && sit && /12X36|DIA SIM/.test(str(r.DESCESCALA).toUpperCase()) && parRE[String(r.RE)] == null) {
        parRE[String(r.RE)] = sit === 'FOLGA' ? 1 - diaNum(d) % 2 : diaNum(d) % 2;
      }
    });
    if (mov) mov.descobertos.forEach(function (x) {
      if (!x.vaga) return;
      descobertaNoDia[x.data + '|' + x.vaga] = 1;
      if (parVaga[x.vaga] == null) parVaga[x.vaga] = diaNum(x.data) % 2;
    });
    var feriasEmFolga = 0;
    function feriasEmDiaDeTrabalho(r, d) {
      var v = String(r.COD_VAGA);
      if (cobertaNoDia[d + '|' + v] || descobertaNoDia[d + '|' + v]) return true;
      var par = parRE[String(r.RE)] != null ? parRE[String(r.RE)] : parVaga[v];
      return trabalhaNaEscala(r.DESCESCALA, d, par) === true;
    }

    var vagaNoQuadro = {};
    rows.forEach(function (r) {
      if (str(r.DESCTPCOBERTURA) !== 'EFETIVO' || vazio(r.COD_VAGA)) return;
      var d = dataStr(r.DATA), sit = situacao(r.DESCSITUACAOHOJE);
      var ferias = !sit && mov && emFerias(mov.feriasPorRE, r.RE, d) && feriasEmDiaDeTrabalho(r, d);
      if (QE_STATUS[sit] || ferias) vagaNoQuadro[d + '|' + String(r.COD_VAGA)] = 1;
    });

    rows.forEach(function (r) {
      var data = dataStr(r.DATA), cliente = str(r.NOMECLIENTE),
          posto = str(r.NOMELOCAL), cargo = str(r.DESC_CARGO);
      // empresa "dona" do posto: prefere a da movimentação (empresa da vaga/contrato); sem isso,
      // usa a do ponto (que é a empresa de QUEM BATEU O PONTO naquele dia — pode ser emprestado
      // de outra empresa numa cobertura, e nesse caso o posto "vazaria" para a empresa errada)
      var emp = (mov && mov.empresaPorLocal[posto]) || str(r.NOMEEMPRESA);
      if (!data || !emp || !posto || !cargo) return;
      if (excluido(cliente, posto)) return;                       // Reserva Técnica / MaxxSeg / Gotham-ADM
      var tn = turno(r.HRENTRADA), tipo = str(r.DESCTPCOBERTURA), sit = situacao(r.DESCSITUACAOHOJE);
      // situação em branco: no ponto isso é ambíguo (férias, afastamento, etc. tudo aparece vazio).
      // Se a movimentação confirma que é férias, conta como quadro esperado (o posto pode continuar
      // precisando de gente), igual a uma falta — sem a movimentação, fica como estava (não conta).
      var emFer = !sit && mov && emFerias(mov.feriasPorRE, r.RE, data);
      var ferias = emFer && feriasEmDiaDeTrabalho(r, data);
      if (ferias) feriasContadas++; else if (emFer) feriasEmFolga++;
      var k = [data, emp, posto, cargo, tn].join('|');
      var g = grupos[k];
      if (!g) { g = grupos[k] = { data: data, emp: emp, posto: posto, cargo: cargo, turno: tn, qe: 0, presentes: 0, reservas: 0, faltantes: [] }; ordem.push(k); }
      if (tipo === 'EFETIVO' && (QE_STATUS[sit] || ferias)) g.qe++;
      if (RESERVA_TIPOS[tipo] && sit === 'TRABALHO' && !vazio(r.COD_VAGA) && !vagaNoQuadro[data + '|' + String(r.COD_VAGA)]) {
        vagaNoQuadro[data + '|' + String(r.COD_VAGA)] = 1;                     // uma vez por vaga e dia
        g.qe++; g.vagaAberta = true; vagasAbertasCobertas++;
      }
      if (sit === 'TRABALHO') { g.presentes++; if (vazio(r.COD_VAGA)) g.presSemVaga = (g.presSemVaga || 0) + 1; }
      if (RESERVA_TIPOS[tipo] && sit === 'TRABALHO') g.reservas++;
      if (tipo === 'EFETIVO' && sit === 'FALTA') g.faltantes.push(str(r.NOMEFUNCIONARIO));
      if (tipo === 'EFETIVO' && ferias) g.faltantes.push(str(r.NOMEFUNCIONARIO) + ' (férias)');
      empresas[emp] = 1;

      // códigos e atributos estáveis por empresa+posto+cargo: guarda o último valor preenchido
      var mk = [emp, posto, cargo].join('|');
      var mt = meta[mk] || (meta[mk] = {});
      ['COD_EMPRESA', 'COD_CLIENTE', 'COD_LOCAL'].concat(EXTRAS).forEach(function (c) {
        if (!vazio(r[c])) mt[c] = r[c];
      });
    });

    // "posto descoberto" registrado na movimentação = dia de trabalho sem ninguém na vaga. Entra como falta,
    // a menos que o ponto já mostre essa vaga no quadro do dia (titular escalado/férias ou vaga ocupada).
    var descobertosIncluidos = 0, datasPonto = {};
    ordem.forEach(function (k) { datasPonto[grupos[k].data] = 1; });
    if (mov) mov.descobertos.forEach(function (x) {
      if (!datasPonto[x.data] || !x.local || !x.cargo || excluido(x.cliente, x.local)) return;
      if (x.vaga && vagaNoQuadro[x.data + '|' + x.vaga]) return;
      if (x.vaga) vagaNoQuadro[x.data + '|' + x.vaga] = 1;
      var emp = mov.empresaPorLocal[x.local] || x.emp;
      var k = [x.data, emp, x.local, x.cargo, x.turno].join('|');
      var g = grupos[k];
      if (!g) { g = grupos[k] = { data: x.data, emp: emp, posto: x.local, cargo: x.cargo, turno: x.turno, qe: 0, presentes: 0, reservas: 0, faltantes: [] }; ordem.push(k); }
      g.qe++; g.descoberto = true; descobertosIncluidos++; empresas[emp] = 1;
    });

    // nível 2: vaga sem titular + cobertura atípica entre cargos do mesmo empresa+posto+turno
    var porPT = {};
    ordem.forEach(function (k) {
      var g = grupos[k];
      g.deficit = Math.max(g.qe - g.presentes, 0);
      g.sobra = Math.max(g.presentes - g.qe, 0);
      g.efet = g.presentes;
      g.faltas = g.deficit;
      g.atipica = false;
      g.vaga = !!g.vagaAberta;
      var kp = [g.data, g.emp, g.posto, g.turno].join('|');
      (porPT[kp] = porPT[kp] || []).push(g);
    });
    Object.keys(porPT).forEach(function (kp) {
      var grp = porPT[kp];
      grp.forEach(function (g) {
        // reserva para presenças SEM código de vaga (a regra por vaga não as alcança): cargo sem nenhum titular
        // escalado e com gente trabalhando sem vaga identificada = vaga sem titular. Presenças com vaga já foram
        // tratadas pela regra por vaga e, se sobrarem, são transferidas como cobertura atípica.
        if (g.qe === 0 && (g.presSemVaga || 0) > 0) {
          g.qe = g.presSemVaga; g.deficit = 0; g.sobra = g.presentes - g.qe; g.faltas = 0; g.vaga = true;
        }
      });
      var sobra = 0, deficit = 0;
      grp.forEach(function (g) { sobra += g.sobra; deficit += g.deficit; });
      if (sobra <= 0 || deficit === 0) return;
      // cobertura atípica entre funções: a presença sobrando num cargo cobre a falta de outro cargo do mesmo
      // posto+turno. Ela é TRANSFERIDA (sai do cargo que sobrou e entra no que faltou), para a mesma pessoa
      // não ser contada duas vezes.
      var doadores = grp.filter(function (g) { return g.sobra > 0; });
      for (var i = 0; i < grp.length && sobra > 0; i++) {
        var g = grp[i];
        if (g.deficit <= 0) continue;
        var usado = Math.min(g.deficit, sobra);
        g.faltas = g.deficit - usado;
        g.efet = g.presentes + usado;
        if (usado > 0) g.atipica = true;
        sobra -= usado;
        for (var k = 0, falta = usado; k < doadores.length && falta > 0; k++) {
          var tira = Math.min(doadores[k].sobra, falta);
          doadores[k].sobra -= tira; doadores[k].efet -= tira; falta -= tira;
          if (tira > 0) doadores[k].atipica = true;
        }
      }
    });

    // nível 3: uma linha por empresa+posto+cargo+dia, com diurno e noturno lado a lado
    var finais = {}, ordemF = [];
    ordem.forEach(function (k) {
      var g = grupos[k];
      var kf = [g.data, g.emp, g.posto, g.cargo].join('|');
      var f = finais[kf];
      if (!f) {
        f = finais[kf] = { data: g.data, emp: g.emp, posto: g.posto, cargo: g.cargo,
          qe_d: 0, qe_n: 0, faltas_d: 0, faltas_n: 0, res_d: 0, res_n: 0, efet_d: 0, efet_n: 0,
          colab: [], atipica: false, vaga: false, descoberto: false };
        ordemF.push(kf);
      }
      var s = g.turno === 'D' ? '_d' : '_n';
      f['qe' + s] += g.qe; f['efet' + s] += g.efet; f['faltas' + s] += g.faltas; f['res' + s] += g.reservas;
      g.faltantes.forEach(function (n) { if (f.colab.indexOf(n) < 0) f.colab.push(n); });
      if (g.atipica) f.atipica = true;
      if (g.vaga) f.vaga = true;
      if (g.descoberto) f.descoberto = true;
    });

    ordemF.sort(function (a, b) {
      var x = finais[a], y = finais[b];
      return (x.data < y.data ? -1 : x.data > y.data ? 1 : 0) ||
             x.emp.localeCompare(y.emp, 'pt-BR') || x.posto.localeCompare(y.posto, 'pt-BR') ||
             x.cargo.localeCompare(y.cargo, 'pt-BR');
    });

    // NOTA: a regra automática de "sem atividade no período" foi removida — ela escondia postos
    // reais descobertos (ex.: único titular em férias sem reserva para cobrir), que são exatamente
    // o tipo de alerta que a planilha deve mostrar, não ocultar. Ver histórico da conversa.
    var semAtividade = [];

    var comObs = 0, semTitular = 0;
    var linhas = ordemF.map(function (kf) {
      var f = finais[kf], mt = meta[[f.emp, f.posto, f.cargo].join('|')] || {};
      var partes = [];
      if (f.atipica) { partes.push('Cobertura atípica entre funções'); }
      if (f.vaga) { partes.push('Vaga sem titular fixo (coberta por reserva/FT)'); semTitular++; }
      if (f.descoberto) partes.push('Posto descoberto (registrado na movimentação)');
      if (partes.length) comObs++;
      var p = f.data.split('-');
      var dia = DIAS[new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay()];
      var pctD = f.qe_d > 0 ? Math.round(f.efet_d / f.qe_d * 10000) / 10000 : '';
      var pctN = f.qe_n > 0 ? Math.round(f.efet_n / f.qe_n * 10000) / 10000 : '';
      function m(c) { return vazio(mt[c]) ? '' : mt[c]; }
      return [f.data, dia, f.emp, m('COD_EMPRESA'), m('COD_CLIENTE'), f.posto, f.cargo,
        f.qe_d, f.qe_n, f.faltas_d, f.faltas_n, f.res_d, f.res_n, f.efet_d, pctD, f.efet_n, pctN,
        f.colab.join('; '), '', partes.join('; '), '',
        m('COD_CLIENTE'), m('COD_LOCAL')].concat(EXTRAS.map(m));
    });

    var datas = {};
    linhas.forEach(function (l) { datas[l[0]] = 1; });
    return { linhas: linhas,
             resumo: { linhas: linhas.length, datas: Object.keys(datas).sort(), empresas: Object.keys(empresas).sort(),
                       comObs: comObs, vagasSemTitular: semTitular,
                       semAtividade: semAtividade.map(function (x) { return x[0] + ' — ' + x[1] + ' (' + x[2] + ')'; }),
                       vagasAbertasCobertas: vagasAbertasCobertas,
                       movimentacao: mov ? { feriasContadas: feriasContadas, feriasEmFolga: feriasEmFolga, descobertosIncluidos: descobertosIncluidos } : null } };
  }

  var api = { agregar: agregar, dataStr: dataStr, turno: turno, parseMovimentacao: parseMovimentacao };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.SAR2G = api;
})(typeof window !== 'undefined' ? window : this);
