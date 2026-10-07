// Sincroniza a planilha do Google (aba EFETIVIDADE) com o D1. Só regrava os dias cujo conteúdo mudou.
// serial do Sheets -> AAAA-MM-DD. Arredonda (não trunca): datas gravadas como "dia anterior, 20:00" por
// diferença de fuso voltam para o dia certo.
const dia = s => new Date(Date.UTC(1899, 11, 30) + Math.round(s) * 864e5).toISOString().slice(0, 10);
const n = v => { const x = Number(v); return Number.isFinite(x) ? Math.trunc(x) : 0; };     // vazio e #DIV/0! viram 0
const t = v => { if (v == null) return null; const s = String(v).trim().replace(/\s+/g, ' '); return s || null; };
const COLS = ['data','dia_semana','empresa','cod_empresa','cod_posto','posto','cargo','seq','qe_d','qe_n','faltas_d','faltas_n','res_d','res_n','efet_d','efet_n','colaborador_falta','reserva_cobertura','obs','obs_extra'];
const MAX_DIAS = 4; // limite de subrequisições por execução; o restante entra na próxima

export function transform(values) {
  const seen = new Map(), days = new Map();
  for (const r of values) {
    if (typeof r[0] !== 'number' || r[0] < 40000) continue;             // linha sem data válida
    // linha-modelo de data futura (quadro, efetividade em branco): ainda não importada, não entra no dashboard
    if ([7, 8, 13, 15].every(i => r[i] === undefined || r[i] === null || r[i] === '')) continue;
    const data = dia(r[0]), cod = n(r[4]), posto = t(r[5]), cargo = t(r[6]), emp = t(r[2]);
    if (!posto || !cargo || !emp) continue;
    const k = [data, cod, posto, cargo].join('|'), seq = seen.get(k) ?? 0; seen.set(k, seq + 1);
    const row = [data, t(r[1]), emp, n(r[3]), cod, posto, cargo, seq, n(r[7]), n(r[8]), n(r[9]), n(r[10]), n(r[11]), n(r[12]), n(r[13]), n(r[15]), t(r[17]), t(r[18]), t(r[19]), t(r[20])];
    if (!days.has(data)) days.set(data, []);
    days.get(data).push(row);
  }
  return days;
}

const sha = async s => [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)))].map(b => b.toString(16).padStart(2, '0')).join('');

export async function sync(env) {
  // lê só até a última linha com dados reais (gravada pelo importador na aba oculta SAR2G_config); as linhas-modelo
  // das datas futuras ficam de fora e a leitura continua leve mesmo com a aba preenchida até o fim do ano
  let faixa = env.SHEET_RANGE || 'EFETIVIDADE!A4:U';
  try {
    const c = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${env.SHEET_ID}/values/${encodeURIComponent('SAR2G_config!B1')}?valueRenderOption=UNFORMATTED_VALUE&key=${env.GOOGLE_API_KEY}`);
    if (c.ok) { const ult = Number(((await c.json()).values || [[]])[0][0]); if (ult >= 4) faixa = faixa.replace(/:([A-Z]+)\d*$/, ':$1' + ult); }
  } catch (e) { /* sem a aba de configuração: lê a faixa inteira (as linhas em branco são ignoradas mesmo assim) */ }
  const range = encodeURIComponent(faixa);
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${env.SHEET_ID}/values/${range}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&key=${env.GOOGLE_API_KEY}`);
  if (!res.ok) throw new Error('Google Sheets ' + res.status + ': ' + (await res.text()).slice(0, 200));
  const days = transform((await res.json()).values || []);
  const old = new Map((await env.DB.prepare('SELECT data, hash FROM sync_state').all()).results.map(r => [r.data, r.hash]));
  const ins = `INSERT INTO efetividade (${COLS.join(',')}) VALUES (${COLS.map(() => '?').join(',')})`;
  let dias = 0, linhas = 0, pendentes = 0;
  for (const [data, rows] of days) {
    const h = await sha(JSON.stringify(rows));
    if (old.get(data) === h) continue;
    if (dias >= MAX_DIAS) { pendentes++; continue; }
    const st = [env.DB.prepare('DELETE FROM efetividade WHERE data = ?').bind(data), ...rows.map(r => env.DB.prepare(ins).bind(...r)),
      env.DB.prepare("INSERT OR REPLACE INTO sync_state (data, hash, atualizado_em) VALUES (?,?,datetime('now'))").bind(data, h)];
    for (let i = 0; i < st.length; i += 100) await env.DB.batch(st.slice(i, i + 100)); // hash só é gravado no último lote
    dias++; linhas += rows.length;
  }
  // dias que estão no banco mas não existem mais na planilha (ex.: datas deslocadas antigas) são removidos.
  // Só roda quando a leitura trouxe dados e não há dias pendentes, para nunca apagar por falha de leitura.
  let removidos = [];
  if (days.size > 0 && pendentes === 0) {
    const noBanco = (await env.DB.prepare('SELECT DISTINCT data FROM efetividade').all()).results.map(r => r.data);
    removidos = noBanco.filter(d => !days.has(d));
    if (removidos.length) await env.DB.batch(removidos.flatMap(d => [
      env.DB.prepare('DELETE FROM efetividade WHERE data = ?').bind(d),
      env.DB.prepare('DELETE FROM sync_state WHERE data = ?').bind(d)]));
  }
  return { dias_na_planilha: days.size, dias_atualizados: dias, linhas_gravadas: linhas, dias_pendentes: pendentes, dias_removidos: removidos };
}
