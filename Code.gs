/**
 * CNG · Controle de Postagens — app web (Google Apps Script)
 * A planilha onde este script está instalado é o banco de dados.
 *
 * >>> TROQUE A SENHA DO GESTOR ABAIXO ANTES DE PUBLICAR <<<
 */
const SENHA_GESTOR = 'troque-esta-senha';

const ABA_REG = 'Registros';
const ABA_CFG = 'Config';
const ABA_PIN = 'PINs';
const COLS = ['ID','Data','Dia','Horário','Atividade','Redes','Tipo','Pauta','Responsável','Link','Status','Observações','Registrado em','Editado por','Criado (ms)'];
const DIAS = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
const ATIVS = ['Postagem','Vídeo (matéria)','Captação de vídeo','Edição de vídeo','Trabalho extra','Outro'];
const PLATS = ['Instagram','TikTok','YouTube','Site'];

const DEFAULT_CFG = {
  weekday: {start: 0, end: 23, step: 1},
  weekend: {start: 0, end: 22, step: 2},
  platCfg: {
    Instagram: {mode: 'grade'},
    TikTok: {mode: 'meta', wd: 8, we: 8},
    YouTube: {mode: 'meta', wd: 4, we: 4},
    Site: {mode: 'meta', wd: 5, we: 3}
  },
  holidays: [
    {date:'2026-11-15',name:'Proclamação da República'},{date:'2026-11-20',name:'Consciência Negra'},
    {date:'2026-12-25',name:'Natal'},{date:'2027-01-01',name:'Confraternização Universal'},
    {date:'2027-02-08',name:'Carnaval'},{date:'2027-02-09',name:'Carnaval'},
    {date:'2027-03-26',name:'Sexta-feira Santa'},{date:'2027-04-21',name:'Tiradentes'},
    {date:'2027-05-01',name:'Dia do Trabalho'},{date:'2027-05-27',name:'Corpus Christi'},
    {date:'2027-09-07',name:'Independência'},{date:'2027-10-12',name:'Nossa Senhora Aparecida'},
    {date:'2027-11-02',name:'Finados'},{date:'2027-11-15',name:'Proclamação da República'},
    {date:'2027-11-20',name:'Consciência Negra'},{date:'2027-12-25',name:'Natal'}
  ],
  team: [
    {name:'Thais', role:'Social media — TikTok, Instagram, Site e YouTube', plats:['Instagram','TikTok','Site','YouTube'], metaVideos:0, metaPosts:0},
    {name:'Iones', role:'Instagram · edição e captação de vídeo', plats:['Instagram'], metaVideos:0, metaPosts:0},
    {name:'Yan',   role:'Instagram e TikTok', plats:['Instagram','TikTok'], metaVideos:0, metaPosts:0},
    {name:'Vitor', role:'Jornalista — 3 vídeos por dia', plats:[], metaVideos:3, metaPosts:0},
    {name:'Bruno', role:'Trabalhos extras', plats:[], metaVideos:0, metaPosts:0}
  ]
};

/* ---------------- web app ---------------- */
function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CNG · Controle de Postagens')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** API usada pelo site no GitHub Pages (o mesmo sistema, servido de fora do Google). */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = {getState: getState, login: login, saveEntry: saveEntry, deleteEntry: deleteEntry, saveConfig: saveConfig, resetPin: resetPin};
    if (!fns[req.fn]) throw new Error('Função inválida.');
    out = {ok: true, result: fns[req.fn].apply(null, req.args || [])};
  } catch (err) {
    out = {ok: false, error: String(err && err.message || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Rode uma vez pelo editor (botão Executar) para criar as abas e autorizar. */
function setup() { ensureSheets_(); }

/* ---------------- planilha ---------------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'America/Sao_Paulo'; }

function ensureSheets_() {
  const ss = ss_();
  let reg = ss.getSheetByName(ABA_REG);
  if (!reg) {
    reg = ss.insertSheet(ABA_REG, 0);
    reg.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#14171c').setFontColor('#ffffff');
    reg.setFrozenRows(1);
    reg.getRange(1, 1, reg.getMaxRows(), COLS.length).setNumberFormat('@');
    reg.setColumnWidth(8, 320);
    reg.hideColumns(1);
    reg.hideColumns(15);
  }
  let cfg = ss.getSheetByName(ABA_CFG);
  if (!cfg) {
    cfg = ss.insertSheet(ABA_CFG);
    cfg.getRange('A1').setNumberFormat('@').setValue(JSON.stringify(DEFAULT_CFG));
    cfg.getRange('A3').setValue('Não edite esta aba à mão: use a tela "Ajustes" do sistema (entrando como Gestor).');
  }
  let pin = ss.getSheetByName(ABA_PIN);
  if (!pin) {
    pin = ss.insertSheet(ABA_PIN);
    pin.getRange(1, 1, 1, 2).setValues([['Nome', 'PIN (codificado)']]).setFontWeight('bold');
    pin.getRange(1, 1, pin.getMaxRows(), 2).setNumberFormat('@');
    pin.hideSheet();
  }
}

function readCfg_() {
  ensureSheets_();
  try {
    const raw = ss_().getSheetByName(ABA_CFG).getRange('A1').getValue();
    const c = JSON.parse(raw);
    return Object.assign(JSON.parse(JSON.stringify(DEFAULT_CFG)), c || {});
  } catch (e) {
    return JSON.parse(JSON.stringify(DEFAULT_CFG));
  }
}

function readPins_() {
  const sh = ss_().getSheetByName(ABA_PIN);
  const n = sh.getLastRow();
  const out = {};
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r, i) {
    if (r[0]) out[String(r[0])] = {hash: String(r[1]), row: i + 2};
  });
  return out;
}

function hash_(s) {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('salt');
  if (!salt) { salt = Utilities.getUuid(); props.setProperty('salt', salt); }
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + s, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

function cellText_(v, kind) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, tz_(), kind === 'date' ? 'yyyy-MM-dd' : 'HH:mm');
  }
  return v == null ? '' : String(v);
}
// a aba mostra a data como dd/mm/aaaa; o sistema usa aaaa-mm-dd
function toIso_(s) {
  const m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : String(s);
}
function toBr_(iso) {
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : String(iso);
}
function dowName_(iso) {
  const d = new Date(iso + 'T12:00:00Z');
  return DIAS[d.getUTCDay()];
}

function unq_(v) { const s = String(v == null ? '' : v); return /^'[=+\-@]/.test(s) ? s.slice(1) : s; }
function rowToEntry_(r) {
  const date = toIso_(cellText_(r[1], 'date'));
  return {
    id: String(r[0]), date: date, slot: cellText_(r[3], 'time'), ativ: String(r[4] || ''),
    plats: String(r[5] || '').split('+').map(function (x) { return x.trim(); }).filter(String),
    tipo: unq_(r[6]), pauta: unq_(r[7]), autor: String(r[8] || ''),
    link: unq_(r[9]), status: unq_(r[10]) || 'Publicado', obs: unq_(r[11]),
    regAt: cellText_(r[12]), editBy: String(r[13] || ''), createdAt: Number(r[14]) || 0
  };
}
function entryToRow_(e) {
  return [e.id, toBr_(e.date), dowName_(e.date), e.slot || '', e.ativ, (e.plats || []).join(' + '), e.tipo,
          e.pauta, e.autor, e.link || '', e.status, e.obs || '', e.regAt || '', e.editBy || '', String(e.createdAt || '')];
}
function readRows_() {
  const sh = ss_().getSheetByName(ABA_REG);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, COLS.length).getValues();
}
function findRow_(id) {
  const rows = readRows_();
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === id) return {idx: i + 2, row: rows[i]};
  return null;
}

/* ---------------- autenticação ---------------- */
function isGestorOk_(token) { return token === hash_('gestor:' + SENHA_GESTOR); }
function auth_(name, token) {
  if (!name || !token) return null;
  if (name === 'Gestor') return isGestorOk_(token) ? {name: name, admin: true} : null;
  const p = readPins_()[name];
  return p && p.hash === token ? {name: name, admin: false} : null;
}
function must_(name, token) {
  const a = auth_(name, token);
  if (!a) throw new Error('SESSAO: faça login de novo.');
  return a;
}

function nowStr_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }

/* ---------------- API chamada pela página ---------------- */
function getState(mk, name, token) {
  ensureSheets_();
  const cfg = readCfg_();
  const pins = readPins_();
  const pinsSet = {};
  Object.keys(pins).forEach(function (k) { pinsSet[k] = true; });
  const a = auth_(name, token);
  const m = String(mk || '').match(/^(\d{4})-(\d{2})$/);
  let prev = '';
  if (m) {
    const y = +m[1], mo = +m[2];
    prev = mo === 1 ? (y - 1) + '-12' : y + '-' + ('0' + (mo - 1)).slice(-2);
  }
  const entries = readRows_().map(rowToEntry_).filter(function (e) {
    const k = e.date.slice(0, 7);
    return e.id && (k === mk || k === prev);
  });
  return {cfg: cfg, pinsSet: pinsSet, entries: entries, auth: a, serverTime: nowStr_()};
}

function login(name, pin) {
  pin = String(pin || '').trim();
  if (name === 'Gestor') {
    if (SENHA_GESTOR === 'troque-esta-senha') return {ok: false, msg: 'Troque a senha do gestor no Code.gs antes de usar.'};
    if (pin !== SENHA_GESTOR) return {ok: false, msg: 'Senha incorreta.'};
    return {ok: true, token: hash_('gestor:' + SENHA_GESTOR), admin: true};
  }
  if (!/^\d{4,6}$/.test(pin)) return {ok: false, msg: 'Use de 4 a 6 números.'};
  const cfg = readCfg_();
  if (!cfg.team.some(function (t) { return t.name === name; })) return {ok: false, msg: 'Funcionário não cadastrado.'};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const pins = readPins_();
    const h = hash_('pin:' + name + ':' + pin);
    if (pins[name]) {
      if (pins[name].hash !== h) return {ok: false, msg: 'PIN incorreto.'};
    } else {
      const sh = ss_().getSheetByName(ABA_PIN);
      sh.appendRow([name, h]);
    }
    return {ok: true, token: h, admin: false, created: !pins[name]};
  } finally { lock.releaseLock(); }
}

function clean_(s, max) { return String(s == null ? '' : s).replace(/[\r\n]+/g, ' ').trim().slice(0, max); }
function cleanEntry_(e) {
  const out = {
    date: clean_(e.date, 10), slot: clean_(e.slot, 5), ativ: clean_(e.ativ, 40),
    plats: (e.plats || []).filter(function (p) { return PLATS.indexOf(p) >= 0; }),
    tipo: clean_(e.tipo, 40) || 'Outro', status: clean_(e.status, 30) || 'Publicado',
    pauta: clean_(e.pauta, 200), link: clean_(e.link, 500), obs: clean_(e.obs, 500)
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(out.date)) throw new Error('Data inválida.');
  if (out.slot && !/^\d{2}:\d{2}$/.test(out.slot)) throw new Error('Horário inválido.');
  if (ATIVS.indexOf(out.ativ) < 0) throw new Error('Atividade inválida.');
  if (!out.pauta) throw new Error('Escreva a pauta.');
  if (out.link && !/^https?:\/\//i.test(out.link)) throw new Error('Link inválido.');
  // evita fórmulas injetadas na planilha
  ['pauta', 'obs', 'tipo', 'link', 'status'].forEach(function (k) { if (/^[=+\-@]/.test(out[k])) out[k] = "'" + out[k]; });
  return out;
}

function saveEntry(name, token, e) {
  const a = must_(name, token);
  const body = cleanEntry_(e || {});
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ss_().getSheetByName(ABA_REG);
    if (e.id) {
      const f = findRow_(String(e.id));
      if (!f) throw new Error('Registro não encontrado (pode ter sido excluído).');
      const old = rowToEntry_(f.row);
      if (!a.admin && old.autor !== a.name) throw new Error('Você só pode editar os seus registros.');
      const full = Object.assign({}, body, {id: old.id, autor: old.autor, regAt: old.regAt, createdAt: old.createdAt, editBy: a.name + ' ' + nowStr_()});
      const rg = sh.getRange(f.idx, 1, 1, COLS.length);
      rg.setNumberFormat('@');
      rg.setValues([entryToRow_(full)]);
      return {ok: true, id: old.id};
    }
    const full = Object.assign({}, body, {id: Utilities.getUuid(), autor: a.name, regAt: nowStr_(), createdAt: Date.now(), editBy: ''});
    const r = sh.getLastRow() + 1;
    const rg = sh.getRange(r, 1, 1, COLS.length);
    rg.setNumberFormat('@');
    rg.setValues([entryToRow_(full)]);
    return {ok: true, id: full.id};
  } finally { lock.releaseLock(); }
}

function deleteEntry(name, token, id) {
  const a = must_(name, token);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const f = findRow_(String(id));
    if (!f) return {ok: true};
    const old = rowToEntry_(f.row);
    if (!a.admin && old.autor !== a.name) throw new Error('Você só pode excluir os seus registros.');
    ss_().getSheetByName(ABA_REG).deleteRow(f.idx);
    return {ok: true};
  } finally { lock.releaseLock(); }
}

function saveConfig(name, token, c) {
  const a = must_(name, token);
  if (!a.admin) throw new Error('Só o gestor altera os ajustes.');
  const num = function (v, d) { return (v === '' || v == null || isNaN(+v)) ? d : +v; };
  const g = function (x, d) { return {start: Math.max(0, Math.min(23, num(x.start, d.start))), end: Math.max(0, Math.min(23, num(x.end, d.end))), step: Math.max(1, Math.min(6, num(x.step, d.step)))}; };
  const out = {
    weekday: g(c.weekday || {}, DEFAULT_CFG.weekday),
    weekend: g(c.weekend || {}, DEFAULT_CFG.weekend),
    platCfg: {},
    holidays: (c.holidays || []).filter(function (h) { return /^\d{4}-\d{2}-\d{2}$/.test(h.date); })
      .map(function (h) { return {date: h.date, name: clean_(h.name, 60) || 'Feriado'}; }),
    team: (c.team || []).filter(function (t) { return clean_(t.name, 40) && t.name !== 'Gestor'; }).map(function (t) {
      return {name: clean_(t.name, 40), role: clean_(t.role, 80), plats: (t.plats || []).filter(function (p) { return PLATS.indexOf(p) >= 0; }),
              metaVideos: Math.max(0, +t.metaVideos || 0), metaPosts: Math.max(0, +t.metaPosts || 0)};
    })
  };
  PLATS.forEach(function (p) {
    const x = (c.platCfg || {})[p] || {};
    const mode = ['grade', 'meta', 'livre'].indexOf(x.mode) >= 0 ? x.mode : 'livre';
    out.platCfg[p] = {mode: mode, wd: Math.max(0, +x.wd || 0), we: Math.max(0, +x.we || 0)};
  });
  ss_().getSheetByName(ABA_CFG).getRange('A1').setNumberFormat('@').setValue(JSON.stringify(out));
  return {ok: true};
}

function resetPin(name, token, target) {
  const a = must_(name, token);
  if (!a.admin) throw new Error('Só o gestor reseta PIN.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = readPins_()[target];
    if (p) ss_().getSheetByName(ABA_PIN).deleteRow(p.row);
    return {ok: true};
  } finally { lock.releaseLock(); }
}
