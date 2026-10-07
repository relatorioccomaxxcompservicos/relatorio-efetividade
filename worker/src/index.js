import { sync } from './sync.js';
// API do dashboard (Cloudflare Worker + D1). Filtros: ?empresa=&posto=&cargo=&de=YYYY-MM-DD&ate=YYYY-MM-DD
const json = (d, s = 200) => new Response(JSON.stringify(d), { status: s, headers: { 'content-type': 'application/json; charset=utf-8', 'access-control-allow-origin': '*', 'cache-control': 'public, max-age=60' } });
const PCT = (n, d) => `ROUND(100.0*SUM(${n})/NULLIF(SUM(${d}),0),2)`;
const AGG = `SUM(qe_d) qe_d, SUM(efet_d) efet_d, ${PCT('efet_d','qe_d')} pct_d, SUM(qe_n) qe_n, SUM(efet_n) efet_n, ${PCT('efet_n','qe_n')} pct_n, SUM(faltas_d+faltas_n) faltas, SUM(res_d+res_n) reservas`;

// Fora do escopo de efetividade operacional (mesma regra do importador SAR2G), aplicado também na API como
// segunda camada de proteção: Reserva Técnica e MaxxSeg (administrativo) de qualquer empresa; Portaria
// Remota e ADM SÓ da Gotham (o "ADM" de outras empresas, como a Hi-Service, continua contando).
const EXCLUSAO = "NOT (posto IN ('RESERVA TÉCNICA','MAXXSEG') OR (empresa = 'GOTHAM' AND posto IN ('PORTARIA REMOTA','ADM')))";

function where(u) {
  const w = [EXCLUSAO], p = [];
  const f = { empresa: 'empresa = ?', posto: 'posto = ?', cargo: 'cargo = ?', de: 'data >= ?', ate: 'data <= ?' };
  for (const k in f) if (u.searchParams.get(k)) { w.push(f[k]); p.push(u.searchParams.get(k)); }
  return ['WHERE ' + w.join(' AND '), p];
}

export default {
  async fetch(req, env) {
    const u = new URL(req.url), [W, P] = where(u), run = async (sql) => (await env.DB.prepare(sql).bind(...P).all()).results;
    if (u.pathname === '/api/sync') { // execução manual: cabeçalho x-sync-token
      if (!env.SYNC_TOKEN || req.headers.get('x-sync-token') !== env.SYNC_TOKEN) return json({ erro: 'não autorizado' }, 401);
      try { return json(await sync(env)); } catch (e) { return json({ erro: String(e) }, 500); }
    }
    try {
      switch (u.pathname) {
        case '/api/filters': {
          const excl = `WHERE ${EXCLUSAO}`;
          return json({
            empresas: (await env.DB.prepare(`SELECT DISTINCT empresa FROM efetividade ${excl} ORDER BY 1`).all()).results.map(r => r.empresa),
            postos: (await env.DB.prepare(`SELECT DISTINCT posto FROM efetividade ${excl} ORDER BY 1`).all()).results.map(r => r.posto),
            cargos: (await env.DB.prepare(`SELECT DISTINCT cargo FROM efetividade ${excl} ORDER BY 1`).all()).results.map(r => r.cargo),
            ...(await env.DB.prepare(`SELECT MIN(data) de, MAX(data) ate FROM efetividade ${excl}`).first()) });
        }
        case '/api/summary': return json((await run(`SELECT ${AGG}, COUNT(DISTINCT posto) postos, COUNT(DISTINCT data) dias FROM efetividade ${W}`))[0]);
        case '/api/daily': return json(await run(`SELECT data, MAX(dia_semana) dia_semana, ${AGG} FROM efetividade ${W} GROUP BY data ORDER BY data`));
        case '/api/ranking': { // ?by=empresa|posto|cargo (padrão: posto) — menor efetividade primeiro
          const by = ['empresa', 'posto', 'cargo'].includes(u.searchParams.get('by')) ? u.searchParams.get('by') : 'posto';
          return json(await run(`SELECT ${by} nome, ${AGG} FROM efetividade ${W} GROUP BY ${by} HAVING SUM(qe_d)>0 ORDER BY pct_d ASC, faltas DESC LIMIT 50`)); }
        case '/api/faltas': return json(await run(`SELECT data, empresa, posto, cargo, faltas_d, faltas_n, res_d, res_n, colaborador_falta, reserva_cobertura FROM efetividade ${W} AND (faltas_d+faltas_n)>0 ORDER BY data DESC, empresa, posto`));
        default: return json({ erro: 'rota não encontrada' }, 404);
      }
    } catch (e) { return json({ erro: String(e) }, 500); }
  },
  async scheduled(ev, env, ctx) { ctx.waitUntil(sync(env)); } // cron: ver wrangler.toml
};
