import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js/+esm';
import { SUPABASE_URL, SUPABASE_ANON_KEY } from './config.js';

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

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

const TABS = { dash: 'Dashboard', lan: 'Lançamentos', con: 'Contratos', rep: 'Corretores', rel: 'Relatório do cliente' };
let D = { lan: [], con: [], rep: [] }, F = {}, tab = 'dash', charts = {}, srcName = 'Google Sheets (Supabase)';

// Parser adaptado para matriz de arrays recebida via Sheets API
function parseRawSheets(sheetsMap) {
  const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const getRows = targetName => {
    const k = Object.keys(sheetsMap).find(s => norm(s) === norm(targetName));
    return k ? sheetsMap[k] : [];
  };

  const tp = t => /^ent/i.test(t || '') ? 'Entrada' : 'Saída';
  const L = [];

  getRows('lancamentos').slice(1).forEach(r => {
    const d = parseDateValue(r[0]);
    const v = Number(r[6]);
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

  return { lan: L, con, rep };
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
const fl = () => D.lan.filter(x => (!F.mes || x.d.startsWith(F.mes)) && (!F.tipo || x.tipo == F.tipo) && (!F.st || x.st == F.st) && (!F.cat || x.cat == F.cat) && (!F.conta || x.conta == F.conta) && (!F.q || (x.desc + x.obs + x.cat).toLowerCase().includes(F.q.toLowerCase())));
const cf = () => D.con.filter(x => (!F.cor || x.cor == F.cor) && (!F.timo || x.tipo == F.timo) && (!F.q || (x.pag + x.imovel).toLowerCase().includes(F.q.toLowerCase())));
const rf = () => D.rep.filter(x => !F.cor || x.cor == F.cor);

function filters() {
  const u = (a, f) => [...new Set(a.map(f).filter(Boolean))].sort();
  const S = (k, l, o, lab = x => x) => `<label>${l}<select data-f="${k}"><option value="">Todos</option>${o.map(x => `<option value="${esc(x)}"${F[k] === x ? ' selected' : ''}>${esc(lab(x))}</option>`).join('')}</select></label>`;
  const Q = `<label>Busca<input data-f="q" placeholder="Digite para buscar" value="${esc(F.q || '')}"></label>`;
  let h = '';
  if (tab == 'dash' || tab == 'lan') h = S('mes', 'Mês', u(D.lan, x => x.d.slice(0, 7)), mn) + S('tipo', 'Tipo', ['Entrada', 'Saída']) + S('st', 'Status', u(D.lan, x => x.st)) + S('cat', 'Categoria', u(tab == 'dash' ? D.lan.filter(OP) : D.lan, x => x.cat)) + S('conta', 'Conta', u(D.lan, x => x.conta)) + (tab == 'lan' ? Q : '');
  if (tab == 'con') h = S('cor', 'Corretor', u(D.con, x => x.cor)) + S('timo', 'Tipo de imóvel', u(D.con, x => x.tipo)) + Q;
  if (tab == 'rep') h = S('cor', 'Corretor', u(D.rep, x => x.cor));
  $('#flt').innerHTML = h + (h ? '<button class="ghost" id="clr">Limpar filtros</button>' : '');
}

document.addEventListener('change', e => { if (e.target.dataset.f) { F[e.target.dataset.f] = e.target.value; body(); } });
document.addEventListener('input', e => { if (e.target.dataset.f == 'q') { F.q = e.target.value; body(); } });
document.addEventListener('click', e => { if (e.target.id == 'clr') { F = {}; filters(); body(); } });

$('#nav').innerHTML = Object.entries(TABS).map(([k, v]) => `<button data-t="${k}">${v}</button>`).join('');
$('#nav').onclick = e => { const t = e.target.dataset.t; if (t) { tab = t; F = {}; filters(); body(); } };

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
    el.innerHTML = sec('Lançamentos realizados', a.filter(x => x.st == 'Realizado'), 'rz', false) + sec('Lançamentos projetados (a realizar)', a.filter(x => x.st != 'Realizado'), 'pj', true) || '<div class="note">Nenhum lançamento encontrado com os filtros atuais.</div>';
  }
  if (tab == 'con') {
    const a = cf(), t = sum(a, x => x.tot), r = sum(a, x => x.rec), i = sum(a, x => x.inad);
    el.innerHTML = kp([['Total contratado', R(t)], ['Recebido', R(r), 'pos'], ['Saldo a receber', R(t - r)], ['Taxa de liquidação', P(t ? r/t : 0)], ['Inadimplente', R(i), i ? 'neg' : '']]) + `<div class="g">${cv('c1', 'Recebido x Pendente por corretor')}</div>` + tbl([['ID'], ['Pagador'], ['Data'], ['Tipo'], ['Imóvel'], ['Corretor'], ['Comissão', 1], ['Recebido', 1], ['Pendente', 1], ['Quitação', 1], ['Inadimpl.', 1]], a.map(x => [x.id, esc(x.pag), fd(x.d), esc(x.tipo), esc(x.imovel), esc(x.cor), R(x.tot), R(x.rec), R(x.pen), P(x.tot ? x.rec/x.tot : 0), x.inad ? `<span class="neg">${R(x.inad)}</span>` : '—']));
    const g = group(a, x => x.cor, x => x.rec), p = group(a, x => x.cor, x => x.pen), K = Object.keys(g);
    mk('c1', { type: 'bar', data: { labels: K, datasets: [{ label: 'Recebido', data: K.map(k => g[k]), backgroundColor: COL[0] }, { label: 'Pendente', data: K.map(k => p[k]), backgroundColor: COL[1] }] }, options: { scales: { x: { stacked: true }, y: { stacked: true } } } });
  }
  if (tab == 'rep') {
    const a = rf(), c = sum(a, x => x.com), p = sum(a, x => x.pago);
    el.innerHTML = kp([['Comissão total', R(c)], ['Repassado', R(p), 'pos'], ['Pendente de repasse', R(c - p), 'neg'], ['Taxa de quitação', P(c ? p/c : 0)]]) + `<div class="g">${cv('c1', 'Repassado x Pendente por corretor')}</div>` + tbl([['ID'], ['Corretor'], ['Data venda'], ['Imóvel'], ['Comissão', 1], ['Repassado', 1], ['Pendente', 1], ['Quitação', 1]], a.map(x => [x.id, esc(x.cor), fd(x.d), esc(x.imovel), R(x.com), R(x.pago), R(x.pen), P(x.com ? x.pago/x.com : 0)]));
    const g = group(a, x => x.cor, x => x.pago), q = group(a, x => x.cor, x => x.com - x.pago), K = Object.keys(g);
    mk('c1', { type: 'bar', data: { labels: K, datasets: [{ label: 'Repassado', data: K.map(k => g[k]), backgroundColor: COL[0] }, { label: 'Pendente', data: K.map(k => Math.max(0, q[k])), backgroundColor: COL[1] }] }, options: { scales: { x: { stacked: true }, y: { stacked: true } } } });
  }
  if (tab == 'rel') el.innerHTML = rel();
}

function rel() {
  const O = D.lan.filter(OP), A = O.filter(x => x.st == 'Realizado'), V = O.filter(x => x.st != 'Realizado'), en = x => x.tipo == 'Entrada';
  const e = sum(A.filter(en), x => x.v), s = sum(A.filter(x => !en(x)), x => x.v), pe = sum(V.filter(en), x => x.v), ps = sum(V.filter(x => !en(x)), x => x.v);
  const fx = sum(A.filter(x => !en(x) && x.grupo == 'Fixo'), x => x.v);
  const ds = A.map(x => x.d).sort(), cs = Object.entries(group(A.filter(x => !en(x)), x => x.cat, x => x.v)).sort((a, b) => b[1] - a[1]);
  const me = Object.entries(group(A.filter(en), x => x.d.slice(0, 7), x => x.v)).sort((a, b) => b[1] - a[1]);
  const t = sum(D.con, x => x.tot), r = sum(D.con, x => x.rec), i = sum(D.con, x => x.inad), c = sum(D.rep, x => x.com), p = sum(D.rep, x => x.pago);
  const byT = Object.entries(group(D.con, x => x.tipo, x => 1)), pend = Object.entries(group(D.rep, x => x.cor, x => x.com - x.pago)).sort((a, b) => b[1] - a[1]);
  const att = [];
  if (i > 0) att.push(`Há ${R(i)} em parcelas em atraso (${P(t ? i/t : 0)} do total contratado). Recomenda-se acompanhar esses clientes de perto.`);
  if (ps > pe && V.length) att.push(`Os próximos meses têm mais despesas previstas (${R(ps)}) do que entradas previstas (${R(pe)}); vale planejar o caixa.`);
  if (c && p/c < .8 && pend.length) att.push(`Ainda restam ${R(c - p)} de comissões a repassar; o maior saldo está com ${esc(pend[0][0])} (${R(pend[0][1])}).`);
  if (cs.length && s && cs[0][1]/s > .3) att.push(`A categoria “${esc(cs[0][0])}” concentra ${P(cs[0][1]/s)} das despesas.`);
  if (!att.length) att.push('Nenhum ponto crítico identificado nos dados importados.');
  return `<div class="rel"><div class="c">
  <h2 style="margin-top:0">Resumo executivo</h2>
  <p class="note">Relatório gerado automaticamente em ${new Date().toLocaleDateString('pt-BR')} a partir de ${esc(srcName)}${ds.length ? ` · período realizado: ${fd(ds[0])} a ${fd(ds[ds.length-1])}` : ''}.</p>
  ${kp([['Entradas realizadas', R(e), 'pos'], ['Despesas realizadas', R(s), 'neg'], ['Resultado', R(e - s), e >= s ? 'pos' : 'neg'], ['A receber em comissões', R(t - r)]])}
  <h2>1. Como está o seu caixa</h2><p class="note">Considera apenas o operacional: receitas de comissões e despesas fixas e variáveis. Retiradas de lucro, aplicações, resgates e receitas financeiras não entram no cálculo.</p>
  <p>No período, a empresa recebeu <b>${R(e)}</b> em receitas e teve <b>${R(s)}</b> em despesas (<b>${R(fx)}</b> fixas e <b>${R(s - fx)}</b> variáveis), fechando com um resultado de <b class="${e >= s ? 'pos' : 'neg'}">${R(e - s)}</b>${e ? ` (${P((e - s)/e)} do que entrou)` : ''}.${me.length ? ` O melhor mês em entradas foi <b>${mn(me[0][0])}</b>, com ${R(me[0][1])}.` : ''}${cs.length ? ` A maior despesa é <b>${esc(cs[0][0])}</b>, com ${R(cs[0][1])}.` : ''}</p>
  ${V.length ? `<p>Para os próximos meses, estão previstos <b>${R(pe)}</b> de entradas e <b>${R(ps)}</b> de despesas.</p>` : ''}
  <h2>2. Vendas e comissões a receber</h2>
  <p>São <b>${D.con.length} contratos</b> acompanhados, somando <b>${R(t)}</b> em comissões. Já foram recebidos <b>${R(r)}</b> (${P(t ? r/t : 0)}), e restam <b>${R(t - r)}</b> a receber em parcelas.${byT.length ? ` Distribuição dos contratos: ${byT.map(([k, v]) => `${esc(k)} (${v})`).join(', ')}.` : ''}</p>
  <h2>3. Repasses aos corretores</h2>
  <p>As comissões dos corretores somam <b>${R(c)}</b>. Foram repassados <b>${R(p)}</b> (${P(c ? p/c : 0)}), restando <b>${R(c - p)}</b> a repassar conforme os recebimentos.</p>
  <h2>4. Pontos de atenção</h2><ul>${att.map(x => `<li>${x}</li>`).join('')}</ul></div></div>`;
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
    $('#auth-screen').classList.add('hidden');
    $('#app-screen').classList.remove('hidden');
    fetchRemoteSheets();
  } else {
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