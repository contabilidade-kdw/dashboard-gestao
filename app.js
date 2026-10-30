// Importa o Supabase diretamente
import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js/+esm'

// Cole suas chaves reais do Supabase aqui
const supabaseUrl = 'https://zdlfmrouqpshhewaqcbg.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpkbGZtcm91cXBzaGhld2FxY2JnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTMyMDYyMDMsImV4cCI6MjEwODc4MjIwM30.z2glHFoaNsrli38Q7m9aoTX7estZQyaL2CosUZZqHlw';

// Inicia o Supabase
const supabase = createClient(supabaseUrl, supabaseKey);

// ... (resto do seu código do app.js continua igual abaixo)

// Utilitários de Formatação
const $ = s => document.querySelector(s);
const R = v => (+v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const P = v => ((v || 0) * 100).toFixed(1).replace('.', ',') + '%';
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);
const group = (a, k, f) => { const m = {}; a.forEach(x => m[k(x)] = (m[k(x)] || 0) + f(x)); return m; };
const fd = d => d ? d.split('-').reverse().join('/') : '';
const mn = m => new Date(m + '-15T12:00').toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' }).replace('.', '');

const parseDateValue = val => {
  if (!val) return '';
  if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val)) return val.slice(0, 10);
  if (typeof val === 'string' && val.includes('/')) {
    const p = val.split('/');
    if (p.length === 3) return `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
  }
  if (typeof val === 'number') {
    // Número serial de data do Excel / Sheets
    const d = new Date((val - 25569) * 86400 * 1000);
    return !isNaN(d) ? d.toISOString().slice(0, 10) : '';
  }
  const d = new Date(val);
  return d instanceof Date && !isNaN(d) ? d.toISOString().slice(0, 10) : '';
};

const COL = ['#0f5c4d', '#c8963e', '#3b7ea1', '#b3382c', '#7a5c99', '#5f8f3c', '#d17b49', '#6b7686'];
Chart.defaults.color = '#8a94a3';
Chart.defaults.maintainAspectRatio = false;

const TABS = { dash: 'Dashboard', lan: 'Lançamentos', con: 'Contratos', ban: 'Bancos', dre: 'DRE', rep: 'Corretores', rel: 'Relatório do cliente' };
let D = { lan: [], con: [], rep: [], sal: [] }, F = {}, tab = 'dash', charts = {}, srcName = 'Google Sheets (Supabase)';
let metas = [];

// Parser adaptado para matriz de arrays recebida via Sheets API
function parseRawSheets(sheetsMap) {
  const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const getRows = targetName => {
    const k = Object.keys(sheetsMap).find(s => norm(s) === norm(targetName));
    return k ? sheetsMap[k] : [];
  };

  const tp = t => /^ent/i.test(t || '') ? 'Entrada' : 'Saída';
  const L = [];
  // Saldo anterior do BB PJ que a leitura normal não aproveita (sem data, valor em texto ou status diferente de Realizado)
  const SAL = [];
  const reSaldo = /saldo\s*(anterior|inicial|de\s*abertura)/i, reBB = /bb\s*pj/i;
  const nv = x => {
    if (typeof x === 'number') return x;
    let t = String(x ?? '').replace(/R\$|\s/g, '');
    if (/^\(.*\)$/.test(t)) t = '-' + t.slice(1, -1);
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    return t === '' ? NaN : Number(t);
  };

  getRows('lancamentos').slice(1).forEach(r => {
    const d = parseDateValue(r[0]);
    const v = Number(r[6]);
    if ([r[1], r[4], r[2]].some(x => reSaldo.test(String(x ?? ''))) && [r[1], r[4], r[8]].some(x => reBB.test(String(x ?? ''))) && !(d && !isNaN(v) && (r[7] || 'Realizado') == 'Realizado')) {
      const sv = nv(r[6]);
      if (!isNaN(sv)) SAL.push({ d, desc: r[1] || 'Saldo anterior', conta: r[8] || 'BB PJ', st: 'Realizado', v: /^sa[ií]da/i.test(String(r[3] ?? '')) ? -Math.abs(sv) : sv });
    }
    if (d && !isNaN(v)) {
      L.push({
        d,
        desc: r[1] || '',
        grupo: r[2] || '',
        tipo: tp(r[3]),
        cat: r[4] || 'Outros',
        obs: r[5] || '',
        v,
        st: r[7] || 'Realizado',
        conta: r[8] || ''
      });
    }
  });

  getRows('lancamentos projetados').slice(2).forEach(r => {
    [0, 10].forEach(o => {
      const d = parseDateValue(r[o]);
      const v = Number(r[o + 4]);
      if (d && !isNaN(v)) {
        L.push({
          d,
          desc: r[o + 1] || '',
          tipo: tp(r[o + 2]),
          cat: r[o + 3] || 'Outros',
          obs: '',
          v,
          st: r[o + 5] || 'Previsto',
          conta: r[o + 6] || ''
        });
      }
    });
  });

  const gm = {};
  L.forEach(x => { if (x.grupo) gm[x.cat.toLowerCase()] = x.grupo; });
  L.forEach(x => { if (!x.grupo) x.grupo = gm[x.cat.toLowerCase()] || (/pr[óo]-?labore/i.test(x.cat) ? 'Fixo' : ''); });

  const tb = (sheetName, need) => {
    const a = getRows(sheetName);
    const i = a.findIndex(r => need.every(x => r.map(c => String(c).toUpperCase()).includes(x)));
    return i < 0 ? [] : a.slice(i + 1).map(r => ({ r, h: a[i].map(c => String(c).toUpperCase()) }));
  };

  const con = tb('controle_comissoes_v2', ['PAGADOR', 'ID'])
    .filter(({ r, h }) => !isNaN(Number(r[h.indexOf('ID')])))
    .map(({ r, h }) => {
      const g = n => r[h.indexOf(n)];
      return {
        id: g('ID'),
        pag: g('PAGADOR') || '',
        d: parseDateValue(g('DATA VENDA')),
        tipo: g('TIPO') || '',
        imovel: g('IMÓVEL') || '',
        tot: Number(g('VALOR TOTAL DA COMISSÃO')) || 0,
        rec: Number(g('VALOR RECEBIDO')) || 0,
        pen: Number(g('SALDO PENDENTE RECEBIMENTO')) || 0,
        inad: Number(g('Valor Inadimplente') || g('VALOR INADIMPLENTE')) || 0,
        cor: g('CORRETOR') || '—'
      };
    });

  const rep = tb('repasses_comissoes', ['CORRETOR', 'VALOR DA COMISSÃO'])
    .filter(({ r, h }) => !isNaN(Number(r[h.indexOf('ID')])))
    .map(({ r, h }) => {
      const g = n => r[h.indexOf(n)];
      return {
        id: g('ID'),
        cor: g('CORRETOR') || '—',
        d: parseDateValue(g('DATA DA VENDA')),
        imovel: g('IMÓVEL VENDIDO') || '',
        com: Number(g('VALOR DA COMISSÃO')) || 0,
        pago: Number(g('TOTAL REPASSADO')) || 0,
        pen: Number(g('SALDO PENDENTE DE REPASSE')) || 0
      };
    });

  return { lan: L, con, rep, sal: SAL };
}

function loadData(d, name) {
  if (!d.lan.length && !d.con.length && !d.rep.length) {
    return toast('Não foram localizados dados nas abas esperadas.');
  }
  D = d;
  srcName = name;
  F = {};
  $('#src').textContent = `Fonte: ${name} · ${d.lan.length} lançamentos · ${d.con.length} contratos · atualizado às ${new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`;
  filters();
  body();
  toast('Dados carregados com sucesso.');
}

function toast(m) {
  const t = $('#toast');
  t.textContent = m;
  t.style.display = 'block';
  setTimeout(() => t.style.display = 'none', 3500);
}

// Fallback de importação manual caso queira arrastar um arquivo XLSX local
async function imp(f) {
  if (!f) return;
  try {
    const ab = await f.arrayBuffer();
    const wb = XLSX.read(ab, { cellDates: true });
    const sheetsMap = {};
    wb.SheetNames.forEach(n => {
      sheetsMap[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
    });
    loadData(parseRawSheets(sheetsMap), f.name);
  } catch (e) {
    toast('Erro ao processar o arquivo.');
  }
}

$('#file').onchange = e => { imp(e.target.files[0]); e.target.value = ''; };
addEventListener('dragover', e => { e.preventDefault(); $('#drop').style.display = 'flex'; });
addEventListener('dragleave', e => { if (!e.relatedTarget) $('#drop').style.display = 'none'; });
addEventListener('drop', e => { e.preventDefault(); $('#drop').style.display = 'none'; imp(e.dataTransfer.files[0]); });

const OP = x => ['Receitas Operacionais', 'Fixo', 'Variável'].includes(x.grupo);
const ct = x => x.conta || 'Sem conta';

// ---------- Filtro de período (presets + "de" / "até") ----------
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const PER = { '': 'Todo o período', mes: 'Mês atual', ant: 'Mês anterior', t3: 'Últimos 3 meses', t6: 'Últimos 6 meses', t12: 'Últimos 12 meses', ano: 'Ano atual', anoant: 'Ano anterior', custom: 'Período específico' };
const perRange = k => {
  const t = new Date(), y = t.getFullYear(), m = t.getMonth();
  const f = (yy, mm) => iso(new Date(yy, mm, 1)), l = (yy, mm) => iso(new Date(yy, mm + 1, 0));
  return ({
    mes: [f(y, m), l(y, m)], ant: [f(y, m - 1), l(y, m - 1)], t3: [f(y, m - 2), l(y, m)], t6: [f(y, m - 5), l(y, m)],
    t12: [f(y, m - 11), l(y, m)], ano: [f(y, 0), l(y, 11)], anoant: [f(y - 1, 0), l(y - 1, 11)]
  })[k] || ['', ''];
};
const inPer = d => (!F.de || (d && d >= F.de)) && (!F.ate || (d && d <= F.ate));
const perHtml = () => `<label>Período<select data-f="per">${Object.entries(PER).map(([k, l]) => `<option value="${k}"${(F.per || '') === k ? ' selected' : ''}>${l}</option>`).join('')}</select></label>` +
  `<label>De<input type="date" data-f="de" value="${esc(F.de || '')}"></label><label>Até<input type="date" data-f="ate" value="${esc(F.ate || '')}"></label>`;

const fl = () => D.lan.filter(x => (!F.mes || x.d.startsWith(F.mes)) && inPer(x.d) && (!F.tipo || x.tipo == F.tipo) && (!F.st || x.st == F.st) && (!F.cat || x.cat == F.cat) && (!F.conta || x.conta == F.conta) && (!F.q || (x.desc + x.obs + x.cat).toLowerCase().includes(F.q.toLowerCase())));
const cf = () => D.con.filter(x => (!F.pag || x.pag == F.pag) && (!F.timo || x.tipo == F.timo) && inPer(x.d) && (!F.q || (x.pag + x.imovel).toLowerCase().includes(F.q.toLowerCase())));
const rf = () => D.rep.filter(x => !F.cor || x.cor == F.cor);

function filters() {
  const u = (a, f) => [...new Set(a.map(f).filter(Boolean))].sort();
  const S = (k, l, o, lab = x => x) => `<label>${l}<select data-f="${k}"><option value="">Todos</option>${o.map(x => `<option value="${esc(x)}"${F[k] === x ? ' selected' : ''}>${esc(lab(x))}</option>`).join('')}</select></label>`;
  const Q = `<label>Busca<input data-f="q" placeholder="Digite para buscar" value="${esc(F.q || '')}"></label>`;
  let h = '';
  if (tab == 'dash' || tab == 'lan') h = S('mes', 'Mês', u(D.lan, x => x.d.slice(0, 7)), mn) + (tab == 'dash' ? perHtml() : '') + S('tipo', 'Tipo', ['Entrada', 'Saída']) + S('st', 'Status', u(D.lan, x => x.st)) + S('cat', 'Categoria', u(tab == 'dash' ? D.lan.filter(OP) : D.lan, x => x.cat)) + S('conta', 'Conta', u(D.lan, x => x.conta)) + (tab == 'lan' ? Q : '');
  if (tab == 'con') h = S('pag', 'Pagador', u(D.con, x => x.pag)) + perHtml() + S('timo', 'Tipo de imóvel', u(D.con, x => x.tipo)) + Q;
  if (tab == 'ban') h = perHtml() + S('conta', 'Conta bancária', u(D.lan, ct)) + S('tipo', 'Movimento', ['Entrada', 'Saída'], x => x + 's') + Q;
  if (tab == 'dre') h = perHtml() + `<label>Base<select data-f="base"><option value="R"${F.base != 'T' ? ' selected' : ''}>Somente realizado</option><option value="T"${F.base == 'T' ? ' selected' : ''}>Realizado + previsto</option></select></label>`;
  if (tab == 'rep') h = S('cor', 'Corretor', u(D.rep, x => x.cor));
  if (tab == 'rel') h = perHtml();
  $('#flt').innerHTML = h + (h ? '<button class="ghost" id="clr">Limpar filtros</button>' : '');
}

document.addEventListener('change', e => {
  const k = e.target.dataset.f;
  if (!k) return;
  F[k] = e.target.value;
  if (k == 'per') {
    if (F.per != 'custom') { const [a, b] = F.per ? perRange(F.per) : ['', '']; F.de = a; F.ate = b; }
    filters();
  }
  if (k == 'de' || k == 'ate') {
    F.per = 'custom';
    const s = document.querySelector('[data-f="per"]');
    if (s) s.value = 'custom';
    if (F.de && F.ate && F.de > F.ate) toast('A data inicial é maior que a data final.');
  }
  body();
});
document.addEventListener('input', e => { if (e.target.dataset.f == 'q') { F.q = e.target.value; body(); } });
document.addEventListener('click', e => {
  if (e.target.id == 'clr') { F = {}; filters(); body(); }
  if (e.target.id == 'mt-add') addMeta();
  if (e.target.dataset.delmeta) delMeta(e.target.dataset.delmeta);
});

$('#nav').innerHTML = Object.entries(TABS).map(([k, v]) => `<button data-t="${k}">${v}</button>`).join('');
$('#nav').onclick = e => { const t = e.target.dataset.t; if (t) { tab = t; F = {}; filters(); body(); if (t == 'rel') loadMetas().then(() => { if (tab == 'rel') body(); }); } };

const kp = a => `<div class="kp">${a.map(([l, v, c]) => `<div class="k"><small>${l}</small><b class="${c || ''}">${v}</b></div>`).join('')}</div>`;
const cv = (id, t) => `<div class="c"><h3>${t}</h3><div class="ch"><canvas id="${id}"></canvas></div></div>`;
const mk = (id, cfg) => { charts[id] = new Chart($('#' + id), cfg); };
const tbl = (h, r, lim = 400) => `<div class="c tw"><table><thead><tr>${h.map(x => `<th class="${x[1] ? 'n' : ''}">${x[0]}</th>`).join('')}</tr></thead><tbody>${r.slice(0, lim).map(c => `<tr>${c.map((x, i) => `<td class="${h[i][1] ? 'n' : ''}">${x}</td>`).join('')}</tr>`).join('')}</tbody></table></div>${r.length > lim ? `<div class="note">Exibindo ${lim} de ${r.length} linhas. Use os filtros para refinar.</div>` : ''}`;

function body() {
  Object.values(charts).forEach(c => c.destroy());
  charts = {};
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.t == tab));
  $('#pdfh').textContent = `Controle de Gestão · ${TABS[tab]} · ${new Date().toLocaleDateString('pt-BR')}`;
  const el = $('#body');

  if (tab == 'dash') {
    const a = fl().filter(OP), E = a.filter(x => x.tipo == 'Entrada'), Sd = a.filter(x => x.tipo != 'Entrada'), e = sum(E, x => x.v), s = sum(Sd, x => x.v);
    el.innerHTML = kp([['Entradas', R(e), 'pos'], ['Saídas', R(s), 'neg'], ['Resultado', R(e - s), e >= s ? 'pos' : 'neg'], ['Lançamentos', a.length]]) + '<div class="note">Considera apenas o operacional: receitas de comissões e despesas fixas e variáveis.</div>' + `<div class="g">${cv('c1', 'Entradas x Saídas por mês')}${cv('c2', 'Resultado acumulado')}</div><div class="g">${cv('c3', 'Saídas por categoria')}${cv('c4', 'Entradas por categoria')}</div>`;
    const me = group(E, x => x.d.slice(0, 7), x => x.v), ms = group(Sd, x => x.d.slice(0, 7), x => x.v), M = [...new Set([...Object.keys(me), ...Object.keys(ms)])].sort(); let ac = 0;
    mk('c1', { type: 'bar', data: { labels: M.map(mn), datasets: [{ label: 'Entradas', data: M.map(m => me[m] || 0), backgroundColor: COL[0] }, { label: 'Saídas', data: M.map(m => ms[m] || 0), backgroundColor: COL[1] }] } });
    mk('c2', { type: 'line', data: { labels: M.map(mn), datasets: [{ label: 'Acumulado', data: M.map(m => ac += (me[m] || 0) - (ms[m] || 0)), borderColor: COL[0], backgroundColor: COL[0] + '33', fill: true, tension: .25 }] }, options: { plugins: { legend: { display: false } } } });
    const dn = (id, L) => { const g = Object.entries(group(L, x => x.cat, x => x.v)).sort((a, b) => b[1] - a[1]).slice(0, 8); mk(id, { type: 'doughnut', data: { labels: g.map(x => x[0]), datasets: [{ data: g.map(x => x[1]), backgroundColor: COL, borderWidth: 0 }] }, options: { plugins: { legend: { position: 'right' } } } }); };
    dn('c3', Sd); dn('c4', E);
  }
  if (tab == 'lan') {
  const a = fl(), sec = (t, L, cls, asc) => {
    if (!L.length) return '';
    L = L.slice().sort((x, y) => asc ? x.d.localeCompare(y.d) : y.d.localeCompare(x.d));
    const e = sum(L.filter(x => x.tipo == 'Entrada'), x => x.v), s = sum(L.filter(x => x.tipo != 'Entrada'), x => x.v);
    return `<div class="sh ${cls}">${t}<span>${L.length} lançamentos</span></div>` + kp([['Entradas', R(e), 'pos'], ['Saídas', R(s), 'neg'], ['Resultado', R(e - s), e >= s ? 'pos' : 'neg']]) + tbl([['Data'], ['Descrição'], ['Categoria'], ['Tipo'], ['Valor', 1], ['Status'], ['Conta']], L.map(x => [fd(x.d), esc(x.desc.slice(0, 70)), esc(x.cat), x.tipo, `<span class="${x.tipo == 'Entrada' ? 'pos' : 'neg'}">${R(x.v)}</span>`, esc(x.st), esc(x.conta)]));
  };

  const isTodos = !F.st || F.st == 'Todos';

  el.innerHTML = (
    isTodos
      ? sec('Todos os lançamentos', a, 'td', false)
      : sec('Lançamentos realizados', a.filter(x => x.st == 'Realizado'), 'rz', false) + 
        sec('Lançamentos projetados (a realizar)', a.filter(x => x.st != 'Realizado'), 'pj', true)
  ) || '<div class="note">Nenhum lançamento encontrado com os filtros atuais.</div>';
}

  if (tab == 'con') {
    const a = cf(), t = sum(a, x => x.tot), r = sum(a, x => x.rec), i = sum(a, x => x.inad), tx = t ? i / t : 0;
    const vw = F.cv || 'ambos', cg = F.cg || 'pag';
    const opt = (k, v, l) => `<option value="${v}"${(F[k] || (k == 'cv' ? 'ambos' : 'pag')) === v ? ' selected' : ''}>${l}</option>`;
    el.innerHTML = kp([['Total contratado', R(t)], ['Recebido', R(r), 'pos'], ['Saldo a receber', R(t - r)], ['Taxa de liquidação', P(t ? r / t : 0)], ['Inadimplente', R(i), i ? 'neg' : ''], ['Taxa de inadimplência', P(tx), tx ? 'neg' : '']]) +
      '<div class="note">Taxa de inadimplência = valor inadimplente ÷ total contratado. O filtro de período considera a data da venda de cada contrato.</div>' +
      `<div class="g"><div class="c"><div class="ch-h"><h3>Recebido x A receber por ${cg == 'mes' ? 'mês da venda' : 'pagador'}</h3><div class="ctl" style="margin:0"><label>Visão<select data-f="cv">${opt('cv', 'ambos', 'Recebido x A receber')}${opt('cv', 'rec', 'Somente recebido')}${opt('cv', 'pen', 'Somente a receber')}</select></label><label>Agrupar por<select data-f="cg">${opt('cg', 'pag', 'Pagador')}${opt('cg', 'mes', 'Mês da venda')}</select></label></div></div><div class="ch"><canvas id="c1"></canvas></div></div></div>` +
      tbl([['ID'], ['Pagador'], ['Data'], ['Tipo'], ['Imóvel'], ['Corretor'], ['Comissão', 1], ['Recebido', 1], ['A receber', 1], ['Quitação', 1], ['Inadimpl.', 1]], a.map(x => [x.id, esc(x.pag), fd(x.d), esc(x.tipo), esc(x.imovel), esc(x.cor), R(x.tot), R(x.rec), R(x.pen), P(x.tot ? x.rec / x.tot : 0), x.inad ? `<span class="neg">${R(x.inad)}</span>` : '—']));
    const key = cg == 'mes' ? x => x.d ? x.d.slice(0, 7) : 'Sem data' : x => x.pag || 'Sem pagador';
    const g = group(a, key, x => x.rec), p = group(a, key, x => x.pen);
    let K = Object.keys(g);
    if (cg == 'mes') K.sort();
    else {
      K.sort((x, y) => (g[y] + p[y]) - (g[x] + p[x]));
      if (K.length > 15) { const rest = K.slice(15); g['Outros'] = sum(rest, k => g[k]); p['Outros'] = sum(rest, k => p[k]); K = [...K.slice(0, 15), 'Outros']; }
    }
    const ds = [];
    if (vw != 'pen') ds.push({ label: 'Recebido', data: K.map(k => g[k]), backgroundColor: COL[0] });
    if (vw != 'rec') ds.push({ label: 'A receber', data: K.map(k => p[k]), backgroundColor: COL[1] });
    mk('c1', { type: 'bar', data: { labels: cg == 'mes' ? K.map(k => k == 'Sem data' ? k : mn(k)) : K, datasets: ds }, options: vw == 'ambos' ? { scales: { x: { stacked: true }, y: { stacked: true } } } : {} });
  }
  if (tab == 'ban') {
    const sg = x => x.tipo == 'Entrada' ? x.v : -x.v, en = x => x.tipo == 'Entrada';
    const R0 = D.lan.filter(x => x.st == 'Realizado' && (!F.conta || ct(x) == F.conta));
    const salBase = (D.sal || []).filter(x => ct(x) !== 'BB PJ').concat([{ conta: 'BB PJ', v: 20823.74 }]);
    const SA = salBase.filter(x => !F.conta || ct(x) == F.conta);  
    const ini = sum(SA, x => x.v) + (F.de ? sum(R0.filter(x => x.d < F.de), sg) : 0);
    let sd = ini;
    const full = R0.filter(x => inPer(x.d)).sort((x, y) => x.d.localeCompare(y.d)).map(x => ({ ...x, sd: sd += sg(x) }));
    const q = (F.q || '').toLowerCase();
    const rows = full.filter(x => (!F.tipo || x.tipo == F.tipo) && (!q || (x.desc + x.obs + x.cat + ct(x)).toLowerCase().includes(q)));
    const e = sum(rows.filter(en), x => x.v), s = sum(rows.filter(x => !en(x)), x => x.v);
    const cts = [...new Set([...R0.map(ct), ...SA.map(ct)])].sort();
    const resumo = cts.map(c => {
      const L = R0.filter(x => ct(x) == c), i0 = sum(SA.filter(x => ct(x) == c), x => x.v) + (F.de ? sum(L.filter(x => x.d < F.de), sg) : 0), P0 = L.filter(x => inPer(x.d));
      const ce = sum(P0.filter(en), x => x.v), cs = sum(P0.filter(x => !en(x)), x => x.v);
      return [esc(c), R(i0), `<span class="pos">${R(ce)}</span>`, `<span class="neg">${R(cs)}</span>`, `<b class="${i0 + ce - cs >= 0 ? 'pos' : 'neg'}">${R(i0 + ce - cs)}</b>`, i0, ce, cs];
    });
    const tI = sum(resumo, x => x[5]), tE = sum(resumo, x => x[6]), tS = sum(resumo, x => x[7]);
    if (resumo.length > 1) resumo.push(['<b>Total</b>', `<b>${R(tI)}</b>`, `<b class="pos">${R(tE)}</b>`, `<b class="neg">${R(tS)}</b>`, `<b>${R(tI + tE - tS)}</b>`]);
    el.innerHTML = kp([['Saldo inicial', R(ini), ini < 0 ? 'neg' : ''], ['Entradas', R(e), 'pos'], ['Saídas', R(s), 'neg'], ['Variação no período', R(e - s), e >= s ? 'pos' : 'neg'], ['Saldo final', R(sd), sd >= 0 ? 'pos' : 'neg']]) +
      `<div class="note">Considera apenas lançamentos com status Realizado (movimentação efetiva) em todas as contas, incluindo itens não operacionais como retiradas e aplicações. O saldo inicial soma o saldo anterior do BB PJ informado na planilha e as movimentações realizadas anteriores à data “De”${F.tipo ? '; saldo inicial e final consideram todas as movimentações, mesmo com o filtro de entradas/saídas ativo' : ''}.</div>` +
      `<div class="sh">Resumo por conta<span>${cts.length} conta(s)</span></div>` + tbl([['Conta'], ['Saldo inicial', 1], ['Entradas', 1], ['Saídas', 1], ['Saldo final', 1]], resumo.map(x => x.slice(0, 5))) +
      `<div class="g">${cv('c1', 'Entradas x Saídas por mês')}</div>` +
      `<div class="sh">Movimentação<span>${rows.length} lançamentos · mais recentes primeiro</span></div>` +
      tbl([['Data'], ['Conta'], ['Descrição'], ['Categoria'], ['Entrada', 1], ['Saída', 1], ['Saldo', 1]], rows.slice().reverse().map(x => [fd(x.d), esc(ct(x)), esc(x.desc.slice(0, 70)), esc(x.cat), en(x) ? `<span class="pos">${R(x.v)}</span>` : '', en(x) ? '' : `<span class="neg">${R(x.v)}</span>`, `<span class="${x.sd >= 0 ? 'pos' : 'neg'}">${R(x.sd)}</span>`]), 500);
    const me = group(rows.filter(en), x => x.d.slice(0, 7), x => x.v), ms = group(rows.filter(x => !en(x)), x => x.d.slice(0, 7), x => x.v), M = [...new Set([...Object.keys(me), ...Object.keys(ms)])].sort();
    mk('c1', { type: 'bar', data: { labels: M.map(mn), datasets: [{ label: 'Entradas', data: M.map(m => me[m] || 0), backgroundColor: COL[0] }, { label: 'Saídas', data: M.map(m => ms[m] || 0), backgroundColor: COL[1] }] } });
  }
  if (tab == 'dre') el.innerHTML = dre();
  if (tab == 'rep') {
    const a = rf(), c = sum(a, x => x.com), p = sum(a, x => x.pago);
    el.innerHTML = kp([['Comissão total', R(c)], ['Repassado', R(p), 'pos'], ['Pendente de repasse', R(c - p), 'neg'], ['Taxa de quitação', P(c ? p/c : 0)]]) + `<div class="g">${cv('c1', 'Repassado x Pendente por corretor')}</div>` + tbl([['ID'], ['Corretor'], ['Data venda'], ['Imóvel'], ['Comissão', 1], ['Repassado', 1], ['Pendente', 1], ['Quitação', 1]], a.map(x => [x.id, esc(x.cor), fd(x.d), esc(x.imovel), R(x.com), R(x.pago), R(x.pen), P(x.com ? x.pago/x.com : 0)]));
    const g = group(a, x => x.cor, x => x.pago), q = group(a, x => x.cor, x => x.com - x.pago), K = Object.keys(g);
    mk('c1', { type: 'bar', data: { labels: K, datasets: [{ label: 'Repassado', data: K.map(k => g[k]), backgroundColor: COL[0] }, { label: 'Pendente', data: K.map(k => Math.max(0, q[k])), backgroundColor: COL[1] }] }, options: { scales: { x: { stacked: true }, y: { stacked: true } } } });
  }
  if (tab == 'rel') el.innerHTML = rel();
}

// ---------- DRE ----------
function dre() {
  const a = D.lan.filter(x => (F.base == 'T' || x.st == 'Realizado') && inPer(x.d));
  if (!a.length) return '<div class="note">Nenhum lançamento encontrado no período selecionado.</div>';
  const sg = x => x.tipo == 'Entrada' ? x.v : -x.v;
  
  // Classificação: grupos operacionais da planilha; demais itens por palavras da categoria
  const cl = x => {
    if (x.grupo == 'Receitas Operacionais') return 'rec';
    if (x.grupo == 'Variável') return 'var';
    if (x.grupo == 'Fixo') return 'fix';
    const t = `${x.cat} ${x.grupo}`;
    // Retiradas e distribuição de lucros agrupadas junto às movimentações patrimoniais (fora do DRE)
    if (/aplica|resgate|transfer|empr[ée]stimo|retirad|distribui|dividend/i.test(t)) return 'pat';
    if (/financ|rendiment|juros|iof/i.test(t)) return 'fin';
    return 'out';
  };
  
  const M = [...new Set(a.map(x => x.d.slice(0, 7)))].sort();
  const T = {}, pat = [];
  a.forEach(x => {
    const c = cl(x), m = x.d.slice(0, 7);
    if (c == 'pat') { pat.push(x); return; }
    T[c] = T[c] || {}; T[c][x.cat] = T[c][x.cat] || {};
    T[c][x.cat][m] = (T[c][x.cat][m] || 0) + sg(x);
  });
  
  const defs = [
    ['rec', '(+) Receita operacional'], ['var', '(–) Custos e despesas variáveis'], ['mc', '(=) Margem de contribuição', ['rec', 'var']],
    ['fix', '(–) Despesas fixas'], ['ro', '(=) Resultado operacional', ['rec', 'var', 'fix']],
    ['fin', '(+/–) Resultado financeiro'], ['out', '(+/–) Outras receitas e despesas não operacionais'],
    ['rl', '(=) Resultado do período', ['rec', 'var', 'fix', 'fin', 'out']]
  ];
  
  const tot = {}, rows = [];
  const RT = () => sum(M, m => tot.rec[m]);
  const row = (cls, lab, v, tone) => {
    const t = sum(M, m => v[m] || 0), rt = RT();
    const c = x => `<td class="n">${x ? `<span class="${tone ? (x < 0 ? 'neg' : 'pos') : (x < 0 ? 'neg' : '')}">${R(x)}</span>` : '—'}</td>`;
    rows.push(`<tr class="${cls}"><td>${esc(lab)}</td>${M.map(m => c(v[m] || 0)).join('')}${c(t)}<td class="n">${rt ? P(t / rt) : '—'}</td></tr>`);
  };
  
  defs.forEach(([c, lab, parts]) => {
    if (parts) { tot[c] = Object.fromEntries(M.map(m => [m, sum(parts, p => tot[p][m] || 0)])); row('t', lab, tot[c], true); return; }
    const cats = Object.entries(T[c] || {});
    tot[c] = Object.fromEntries(M.map(m => [m, sum(cats, ([, o]) => o[m] || 0)]));
    if (!cats.length && !['rec', 'var', 'fix'].includes(c)) return;
    row('h', lab, tot[c]);
    cats.sort((x, y) => Math.abs(sum(Object.values(y[1]), v => v)) - Math.abs(sum(Object.values(x[1]), v => v))).forEach(([k, o]) => row('i', k, o));
  });
  
  const rt = RT(), ro = sum(M, m => tot.ro[m]), rl = sum(M, m => tot.rl[m]);
  const pn = sum(pat, sg);
  
  return kp([['Receita operacional', R(rt), 'pos'], ['Resultado operacional', R(ro), ro >= 0 ? 'pos' : 'neg'], ['Margem operacional', P(rt ? ro / rt : 0), ro >= 0 ? 'pos' : 'neg'], ['Resultado do período', R(rl), rl >= 0 ? 'pos' : 'neg']]) +
    `<div class="note">DRE gerado a partir dos lançamentos ${F.base == 'T' ? 'realizados e previstos' : 'realizados'} do período, pela data do lançamento. Receitas operacionais, custos variáveis e despesas fixas seguem o grupo da planilha; resultado financeiro e demais itens são identificados pelo nome da categoria. Valores negativos são deduções.${pat.length ? ` Ficam fora do DRE ${pat.length} movimentações patrimoniais (aplicações, resgates, transferências, empréstimos e distribuição de lucros), com saldo líquido de ${R(pn)}.` : ''}</div>` +
    `<div class="c tw"><table class="dre"><thead><tr><th>Descrição</th>${M.map(m => `<th class="n">${mn(m)}</th>`).join('')}<th class="n">Total</th><th class="n">% s/ receita</th></tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
}

// ---------- Metas (compartilhadas via tabela "metas" no Supabase) ----------
const loadMetas = async () => {
  const { data, error } = await supabase.from('metas').select('*').order('created_at', { ascending: true });
  if (error) { console.error('Metas:', error); toast('Não foi possível carregar as metas (a tabela "metas" foi criada no Supabase?).'); return; }
  metas = data || [];
};
const num = s => { s = String(s ?? '').trim(); if (!s) return NaN; if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.'); return Number(s); };
const META_T = { rec: 'Receitas operacionais', res: 'Resultado operacional (entradas − saídas)', desp: 'Limite de despesas operacionais' };

async function addMeta() {
  const nome = $('#mt-nome').value.trim() || 'Meta', tipo = $('#mt-tipo').value, valor = num($('#mt-val').value), de = $('#mt-de').value, ate = $('#mt-ate').value;
  if (!(valor > 0)) return toast('Informe o valor da meta (maior que zero).');
  if (!de || !ate) return toast('Informe o período da meta (de / até).');
  if (de > ate) return toast('A data inicial da meta é maior que a final.');
  const { error } = await supabase.from('metas').insert({ nome, tipo, valor, de, ate });
  if (error) { console.error('Metas:', error); return toast('Não foi possível salvar a meta.'); }
  await loadMetas();
  body();
  toast('Meta adicionada e compartilhada.');
}

async function delMeta(id) {
  const { error } = await supabase.from('metas').delete().eq('id', id);
  if (error) { console.error('Metas:', error); return toast('Não foi possível excluir a meta.'); }
  await loadMetas();
  body();
}

function metaCalc(m) {
  const en = x => x.tipo == 'Entrada', L = D.lan.filter(x => OP(x) && x.d >= m.de && x.d <= m.ate);
  const f = rows => m.tipo == 'rec' ? sum(rows.filter(en), x => x.v) : m.tipo == 'desp' ? sum(rows.filter(x => !en(x)), x => x.v) : sum(rows.filter(en), x => x.v) - sum(rows.filter(x => !en(x)), x => x.v);
  const real = f(L.filter(x => x.st == 'Realizado')), prev = f(L.filter(x => x.st != 'Realizado'));
  return { real, prev, falta: m.valor - real, pct: m.valor ? real / m.valor : 0 };
}

function metasHtml() {
  const t = new Date(), y = t.getFullYear(), mo = t.getMonth();
  const rowsM = metas.map(m => {
    const c = metaCalc(m), desp = m.tipo == 'desp', over = desp ? c.pct > 1 : false;
    const st = desp ? (c.falta >= 0 ? `<span class="pos">Disponível ${R(c.falta)}</span>` : `<span class="neg">Excedido em ${R(-c.falta)}</span>`) : (c.falta > 0 ? R(c.falta) : '<span class="pos">Meta atingida</span>');
    return [`<b>${esc(m.nome)}</b><br><span class="sub">${META_T[m.tipo]}</span>`, `${fd(m.de)} a ${fd(m.ate)}`, R(m.valor), R(c.real), st, `${P(c.pct)}<div class="bar"><i class="${over ? 'over' : ''}" style="width:${Math.max(0, Math.min(100, c.pct * 100))}%"></i></div>`, R(c.prev), R(c.real + c.prev), `<button class="ghost noprint" data-delmeta="${m.id}">Excluir</button>`];
  });
  return `<h2>7. Metas</h2><p class="note">Cadastre uma meta para um período. Os valores são recalculados a cada atualização da planilha, com base nos lançamentos operacionais (receitas de comissões e despesas fixas e variáveis). As metas são compartilhadas entre todos os usuários autorizados.</p>` +
    `<div class="ctl noprint"><label>Nome da meta<input id="mt-nome" placeholder="Ex.: Receita do 1º semestre"></label><label>Tipo<select id="mt-tipo"><option value="rec">${META_T.rec}</option><option value="res">${META_T.res}</option><option value="desp">${META_T.desp}</option></select></label><label>Valor da meta (R$)<input id="mt-val" placeholder="Ex.: 50.000,00"></label><label>De<input type="date" id="mt-de" value="${iso(new Date(y, mo, 1))}"></label><label>Até<input type="date" id="mt-ate" value="${iso(new Date(y, mo + 1, 0))}"></label><button id="mt-add">Adicionar meta</button></div>` +
    (rowsM.length ? tbl([['Meta'], ['Período'], ['Meta (R$)', 1], ['Realizado', 1], ['Falta realizar', 1], ['Progresso', 1], ['Previsto no período', 1], ['Realizado + previsto', 1], ['']], rowsM) : '<div class="note">Nenhuma meta cadastrada ainda.</div>');
}

// ---------- Relatório do cliente ----------
function rel() {
  const LP = D.lan.filter(x => inPer(x.d)), CP = D.con.filter(x => inPer(x.d)), RP = D.rep.filter(x => inPer(x.d));
  const O = LP.filter(OP), A = O.filter(x => x.st == 'Realizado'), V = O.filter(x => x.st != 'Realizado'), en = x => x.tipo == 'Entrada';
  const e = sum(A.filter(en), x => x.v), s = sum(A.filter(x => !en(x)), x => x.v), pe = sum(V.filter(en), x => x.v), ps = sum(V.filter(x => !en(x)), x => x.v);
  const fx = sum(A.filter(x => !en(x) && x.grupo == 'Fixo'), x => x.v);
  const ds = A.map(x => x.d).sort(), cs = Object.entries(group(A.filter(x => !en(x)), x => x.cat, x => x.v)).sort((a, b) => b[1] - a[1]);
  const me = Object.entries(group(A.filter(en), x => x.d.slice(0, 7), x => x.v)).sort((a, b) => b[1] - a[1]);
  const t = sum(CP, x => x.tot), r = sum(CP, x => x.rec), i = sum(CP, x => x.inad), c = sum(RP, x => x.com), p = sum(RP, x => x.pago);
  const byT = Object.entries(group(CP, x => x.tipo, x => 1)), pend = Object.entries(group(RP, x => x.cor, x => x.com - x.pago)).sort((a, b) => b[1] - a[1]);
  const att = [];
  if (i > 0) att.push(`Há ${R(i)} em parcelas em atraso (${P(t ? i / t : 0)} do total contratado). Recomenda-se acompanhar esses clientes de perto.`);
  if (ps > pe && V.length) att.push(`Os próximos meses têm mais despesas previstas (${R(ps)}) do que entradas previstas (${R(pe)}); vale planejar o caixa.`);
  if (c && p / c < .8 && pend.length) att.push(`Ainda restam ${R(c - p)} de comissões a repassar; o maior saldo está com ${esc(pend[0][0])} (${R(pend[0][1])}).`);
  if (cs.length && s && cs[0][1] / s > .3) att.push(`A categoria “${esc(cs[0][0])}” concentra ${P(cs[0][1] / s)} das despesas.`);
  if (!att.length) att.push('Nenhum ponto crítico identificado nos dados importados.');

  // Projeções
  const addM = (ym, n) => { const [y, m] = ym.split('-').map(Number), d = new Date(y, m - 1 + n, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; };
  const last = ds.length ? ds[ds.length - 1].slice(0, 7) : iso(new Date()).slice(0, 7);
  const nMes = ds.length ? (+last.slice(0, 4) - +ds[0].slice(0, 4)) * 12 + (+last.slice(5) - +ds[0].slice(5, 7)) + 1 : 0;
  const avgE = nMes ? e / nMes : 0, avgS = nMes ? s / nMes : 0;
  const rg = +(F.rg ?? 5), rn = +(F.rm ?? 6), cr = +(F.cr ?? 5), cn = +(F.cm ?? 6);
  const pc = o => String(o).replace('.', ',') + '%';
  const so = (k, lab, opts, fmt, def) => `<label>${lab}<select data-f="${k}">${opts.map(o => `<option value="${o}"${+(F[k] ?? def) === o ? ' selected' : ''}>${fmt(o)}</option>`).join('')}</select></label>`;
  const rates = [0, 1, 2, 3, 4, 5, 7.5, 10, 12, 15, 20, 25, 30], mons = [1, 2, 3, 4, 6, 9, 12, 18, 24, 36];
  const proj = (base, rate, n) => Array.from({ length: n }, (_, k) => base * Math.pow(1 + rate / 100, k + 1));
  const rv = proj(avgE, rg, rn), cp = proj(avgS, -cr, cn);
  const tRv = sum(rv, x => x), tCp = sum(cp, x => x);
  const mf = m => `${mn(m)}`;
  const projRec = nMes ? kp([['Média mensal de receitas (base)', R(avgE)], [`Receita estimada em ${rn} meses`, R(tRv), 'pos'], ['Sem aumento (média × meses)', R(avgE * rn)], ['Receita adicional estimada', R(tRv - avgE * rn), tRv - avgE * rn >= 0 ? 'pos' : 'neg']]) +
    tbl([['Mês'], ['Média atual', 1], ['Receita estimada', 1], ['Diferença', 1]], rv.map((v, k) => [mf(addM(last, k + 1)), R(avgE), `<b>${R(v)}</b>`, `<span class="${v - avgE >= 0 ? 'pos' : 'neg'}">${R(v - avgE)}</span>`])) : '<div class="note">Sem receitas realizadas no período para servir de base.</div>';
  const projCus = nMes ? kp([['Média mensal de custos (base)', R(avgS)], [`Custo estimado em ${cn} meses`, R(tCp), 'neg'], ['Sem redução (média × meses)', R(avgS * cn)], ['Economia estimada', R(avgS * cn - tCp), avgS * cn - tCp >= 0 ? 'pos' : 'neg']]) +
    tbl([['Mês'], ['Média atual', 1], ['Custo estimado', 1], ['Economia', 1]], cp.map((v, k) => [mf(addM(last, k + 1)), R(avgS), `<b>${R(v)}</b>`, `<span class="${avgS - v >= 0 ? 'pos' : 'neg'}">${R(avgS - v)}</span>`])) : '<div class="note">Sem despesas realizadas no período para servir de base.</div>';
  const comb = nMes && rn == cn ? `<p>Com as duas premissas por ${rn} meses, o resultado estimado é de <b class="${tRv - tCp >= 0 ? 'pos' : 'neg'}">${R(tRv - tCp)}</b>, contra ${R((avgE - avgS) * rn)} se nada mudar.</p>` : '';

  return `<div class="rel"><div class="c">
  <h2 style="margin-top:0">Resumo executivo</h2>
  <p class="note">Relatório gerado automaticamente em ${new Date().toLocaleDateString('pt-BR')} a partir de ${esc(srcName)}${F.de || F.ate ? ` · filtro de período: ${F.de ? fd(F.de) : 'início'} a ${F.ate ? fd(F.ate) : 'hoje'}` : ''}${ds.length ? ` · período realizado: ${fd(ds[0])} a ${fd(ds[ds.length - 1])}` : ''}.</p>
  ${kp([['Entradas realizadas', R(e), 'pos'], ['Despesas realizadas', R(s), 'neg'], ['Resultado', R(e - s), e >= s ? 'pos' : 'neg'], ['A receber em comissões', R(t - r)]])}
  <h2>1. Como está o seu caixa</h2><p class="note">Considera apenas o operacional: receitas de comissões e despesas fixas e variáveis. Retiradas de lucro, aplicações, resgates e receitas financeiras não entram no cálculo.</p>
  <p>No período, a empresa recebeu <b>${R(e)}</b> em receitas e teve <b>${R(s)}</b> em despesas (<b>${R(fx)}</b> fixas e <b>${R(s - fx)}</b> variáveis), fechando com um resultado de <b class="${e >= s ? 'pos' : 'neg'}">${R(e - s)}</b>${e ? ` (${P((e - s) / e)} do que entrou)` : ''}.${me.length ? ` O melhor mês em entradas foi <b>${mn(me[0][0])}</b>, com ${R(me[0][1])}.` : ''}${cs.length ? ` A maior despesa é <b>${esc(cs[0][0])}</b>, com ${R(cs[0][1])}.` : ''}</p>
  ${V.length ? `<p>Para os próximos meses, estão previstos <b>${R(pe)}</b> de entradas e <b>${R(ps)}</b> de despesas.</p>` : ''}
  <h2>2. Vendas e comissões a receber</h2>
  <p>São <b>${CP.length} contratos</b> acompanhados, somando <b>${R(t)}</b> em comissões. Já foram recebidos <b>${R(r)}</b> (${P(t ? r / t : 0)}), e restam <b>${R(t - r)}</b> a receber em parcelas.${byT.length ? ` Distribuição dos contratos: ${byT.map(([k, v]) => `${esc(k)} (${v})`).join(', ')}.` : ''}</p>
  <h2>3. Repasses aos corretores</h2>
  <p>As comissões dos corretores somam <b>${R(c)}</b>. Foram repassados <b>${R(p)}</b> (${P(c ? p / c : 0)}), restando <b>${R(c - p)}</b> a repassar conforme os recebimentos.</p>
  <h2>4. Pontos de atenção</h2><ul>${att.map(x => `<li>${x}</li>`).join('')}</ul>
  <h2>5. Estimativa de receitas</h2>
  <p class="note">Projeção a partir da média mensal de receitas realizadas no período (${nMes} ${nMes == 1 ? 'mês' : 'meses'}), com aumento composto ao mês, começando em ${ds.length ? mn(addM(last, 1)) : 'o mês seguinte'}.</p>
  <div class="ctl">${so('rg', 'Taxa de aumento de receitas (ao mês)', rates, pc, 5)}${so('rm', 'Quantidade de meses prevista', mons, o => o + (o == 1 ? ' mês' : ' meses'), 6)}</div>
  ${projRec}
  <h2>6. Estimativa de custos</h2>
  <p class="note">Projeção a partir da média mensal de despesas realizadas no período, com redução composta ao mês.</p>
  <div class="ctl">${so('cr', 'Taxa de redução de custos (ao mês)', rates, pc, 5)}${so('cm', 'Quantidade de meses prevista', mons, o => o + (o == 1 ? ' mês' : ' meses'), 6)}</div>
  ${projCus}
  ${comb}
  ${metasHtml()}</div></div>`;
}

// Exportação em PDF
async function pdf() {
  toast('Gerando PDF…');
  document.body.classList.add('pdf');
  try {
    const c = await html2canvas($('#main'), { scale: 2, backgroundColor: '#ffffff' });
    document.body.classList.remove('pdf');
    const doc = new jspdf.jsPDF('p', 'mm', 'a4');
    const W = 210, H = 297, ih = c.height * W / c.width;
    const img = c.toDataURL('image/jpeg', 0.92);
    let left = ih, pos = 0;
    doc.addImage(img, 'JPEG', 0, 0, W, ih);
    left -= H;
    while (left > 0) {
      pos = left - ih;
      doc.addPage();
      doc.addImage(img, 'JPEG', 0, pos, W, ih);
      left -= H;
    }
    doc.save(`controle-de-gestao-${tab}.pdf`);
    toast('PDF gerado.');
  } catch (e) {
    document.body.classList.remove('pdf');
    toast('Não foi possível gerar o PDF.');
  }
}
$('#pdf').onclick = pdf;

// Chamada à Edge Function Supabase
async function fetchRemoteSheets() {
  $('#src').textContent = 'Consultando Supabase Edge Function...';
  const { data, error } = await supabase.functions.invoke('fetch-google-sheets');
  if (error || !data) {
    console.error('Falha na requisição:', error);
    toast('Erro ao buscar dados remotos.');
    $('#src').textContent = 'Erro ao sincronizar com a planilha.';
    return;
  }
  const parsed = parseRawSheets(data);
  loadData(parsed, 'Google Sheets');
}

// Gestão de Autenticação Supabase via E-mail e Senha
supabase.auth.onAuthStateChange((event, session) => {
  if (session) {
    loadMetas().then(() => { if (tab == 'rel') body(); });
    $('#auth-screen').classList.add('hidden');
    $('#app-screen').classList.remove('hidden');
    fetchRemoteSheets();
  } else {
    metas = [];
    $('#auth-screen').classList.remove('hidden');
    $('#app-screen').classList.add('hidden');
    $('#auth-email').value = '';
    $('#auth-pass').value = '';
    $('#auth-msg').textContent = '';
  }
});

// Ação de Login com E-mail e Senha
$('#btn-login').onclick = async () => {
  const email = $('#auth-email').value.trim();
  const password = $('#auth-pass').value;
  const msgEl = $('#auth-msg');

  if (!email || !password) {
    msgEl.textContent = 'Preencha o e-mail e a senha.';
    return;
  }

  msgEl.style.color = 'var(--mut)';
  msgEl.textContent = 'Autenticando...';

  const { data, error } = await supabase.auth.signInWithPassword({
    email: email,
    password: password
  });

  if (error) {
    msgEl.style.color = 'var(--neg)';
    msgEl.textContent = 'E-mail ou senha incorretos.';
  } else {
    msgEl.textContent = '';
  }
};

// Permite acionar o login pressionando Enter no campo de senha
$('#auth-pass').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    $('#btn-login').click();
  }
});

// Logout
$('#btn-logout').onclick = async () => {
  await supabase.auth.signOut();
};
