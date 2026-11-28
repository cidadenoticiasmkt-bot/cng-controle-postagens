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
const ABA_ACT = 'Acoes';
const ACOLS = ['ID','Título','Descrição','Responsável','Prazo','Status','Origem','Criado em','Criado por','Atualizado em'];
const ASTATUS = ['A fazer','Em andamento','Concluído'];
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
    const fns = {getState: getState, login: login, saveEntry: saveEntry, deleteEntry: deleteEntry, saveConfig: saveConfig, resetPin: resetPin, getActions: getActions, saveAction: saveAction, deleteAction: deleteAction,
                 getMetrics: getMetrics, setWindsorKey: setWindsorKey, refreshMetrics: refreshMetrics, sendTestAlert: sendTestAlert, sendTestReport: sendTestReport, automationStatus: automationStatus};
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
              metaVideos: Math.max(0, +t.metaVideos || 0), metaPosts: Math.max(0, +t.metaPosts || 0),
              email: /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(clean_(t.email, 100)) ? clean_(t.email, 100) : ''};
    }),
    alerts: (function (a) { a = a || {};
      const h = function (v, d) { v = +v; return isNaN(v) ? d : Math.max(0, Math.min(23, v)); };
      return {on: !!a.on, from: h(a.from, 7), to: h(a.to, 23), team: a.team !== false, weekly: a.weekly !== false,
              gestor: String(a.gestor || '').split(/[,;\s]+/).filter(function (e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); }).slice(0, 5).join(', ')};
    })(c.alerts)
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


/* ---------------- plano de ação ---------------- */
function actSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(ABA_ACT);
  if (!sh) {
    sh = ss.insertSheet(ABA_ACT);
    sh.getRange(1, 1, 1, ACOLS.length).setValues([ACOLS]).setFontWeight('bold').setBackground('#14171c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), ACOLS.length).setNumberFormat('@');
    sh.hideColumns(1);
  }
  return sh;
}
function actRows_() {
  const sh = actSheet_(); const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, ACOLS.length).getValues().map(function (r, i) {
    return {row: i + 2, id: String(r[0]), titulo: unq_(r[1]), desc: unq_(r[2]), resp: String(r[3] || ''), prazo: toIso_(cellText_(r[4], 'date')),
            status: String(r[5] || 'A fazer'), origem: unq_(r[6]), criadoEm: cellText_(r[7]), criadoPor: String(r[8] || ''), atualizado: cellText_(r[9])};
  }).filter(function (a) { return a.id; });
}
function getActions(name, token) {
  must_(name, token);
  return actRows_().map(function (a) { delete a.row; return a; });
}
function saveAction(name, token, a) {
  const u = must_(name, token);
  a = a || {};
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = actSheet_();
    const status = ASTATUS.indexOf(a.status) >= 0 ? a.status : 'A fazer';
    const safe = function (v, m) { v = clean_(v, m); return /^[=+\-@]/.test(v) ? "'" + v : v; };
    if (a.id) {
      const cur = actRows_().filter(function (x) { return x.id === String(a.id); })[0];
      if (!cur) throw new Error('Ação não encontrada.');
      if (!u.admin) {
        if (cur.resp !== u.name) throw new Error('Só o gestor ou o responsável altera esta ação.');
        sh.getRange(cur.row, 6, 1, 1).setValue(status);
        sh.getRange(cur.row, 10, 1, 1).setValue(u.name + ' ' + nowStr_());
        return {ok: true, id: cur.id};
      }
      const row = [cur.id, safe(a.titulo, 120), safe(a.desc, 600), clean_(a.resp, 40), toBr_(clean_(a.prazo, 10)), status, safe(a.origem, 200), cur.criadoEm, cur.criadoPor, u.name + ' ' + nowStr_()];
      const rg = sh.getRange(cur.row, 1, 1, ACOLS.length); rg.setNumberFormat('@'); rg.setValues([row]);
      return {ok: true, id: cur.id};
    }
    if (!u.admin) throw new Error('Só o gestor cria ações.');
    if (!clean_(a.titulo, 120)) throw new Error('Escreva o título da ação.');
    const id = Utilities.getUuid();
    const row = [id, safe(a.titulo, 120), safe(a.desc, 600), clean_(a.resp, 40), toBr_(clean_(a.prazo, 10)), status, safe(a.origem, 200), nowStr_(), u.name, ''];
    const r = sh.getLastRow() + 1; const rg = sh.getRange(r, 1, 1, ACOLS.length); rg.setNumberFormat('@'); rg.setValues([row]);
    return {ok: true, id: id};
  } finally { lock.releaseLock(); }
}
function deleteAction(name, token, id) {
  const u = must_(name, token);
  if (!u.admin) throw new Error('Só o gestor exclui ações.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const cur = actRows_().filter(function (x) { return x.id === String(id); })[0];
    if (cur) actSheet_().deleteRow(cur.row);
    return {ok: true};
  } finally { lock.releaseLock(); }
}


/* =====================================================================
 * AUTOMAÇÕES: alerta de horário vazio, relatório semanal e métricas
 * Depois de colar este código, rode UMA VEZ a função instalarAutomacoes
 * (menu de funções no topo → instalarAutomacoes → Executar) e autorize.
 * ===================================================================== */
const SITE_URL = 'https://cidadenoticiasmkt-bot.github.io/cng-controle-postagens/';
const ABA_MET = 'Metricas';

function instalarAutomacoes() {
  const handlers = ['verificarHorariosVazios', 'enviarRelatorioSemanal', 'atualizarMetricas'];
  ScriptApp.getProjectTriggers().forEach(function (t) { if (handlers.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('verificarHorariosVazios').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('enviarRelatorioSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(30).inTimezone(tz_()).create();
  ScriptApp.newTrigger('atualizarMetricas').timeBased().everyHours(6).create();
  PropertiesService.getScriptProperties().setProperty('auto_installed', nowStr_());
  Logger.log('Automações instaladas: alerta a cada 15 min, relatório toda segunda 7h30, métricas a cada 6h.');
}

function automationStatus(name, token) {
  const u = must_(name, token); if (!u.admin) throw new Error('Só o gestor.');
  const pr = PropertiesService.getScriptProperties();
  const trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  return {installed: trig.indexOf('verificarHorariosVazios') >= 0, triggers: trig, installedAt: pr.getProperty('auto_installed') || '',
          windsor: !!pr.getProperty('WINDSOR_KEY'), metricsAt: pr.getProperty('metrics_at') || '', metricsErr: pr.getProperty('metrics_err') || '',
          lastAlert: pr.getProperty('last_alert') || '', lastReport: pr.getProperty('last_report') || ''};
}

/* ---------- utilidades de data no fuso da planilha ---------- */
function todayIso_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function addDaysIso_(iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function dowIso_(iso) { return new Date(iso + 'T12:00:00Z').getUTCDay(); }
function isReduced_(cfg, iso) { const w = dowIso_(iso); return w === 0 || w === 6 || (cfg.holidays || []).some(function (h) { return h.date === iso; }); }
function slotsFor_(cfg, iso) { const g = isReduced_(cfg, iso) ? cfg.weekend : cfg.weekday; const step = Math.max(1, +g.step || (isReduced_(cfg, iso) ? 2 : 1));
  const out = []; for (let h = 0; h <= 23; h += step) out.push(('0' + h).slice(-2) + ':00'); return out; }
function gradePlats_(cfg) { return PLATS.filter(function (p) { return ((cfg.platCfg || {})[p] || {}).mode === 'grade'; }); }
function metaPlats_(cfg) { return PLATS.filter(function (p) { return ((cfg.platCfg || {})[p] || {}).mode === 'meta'; }); }
function quota_(cfg, p, iso) { const c = (cfg.platCfg || {})[p] || {}; return isReduced_(cfg, iso) ? (+c.we || 0) : (+c.wd || 0); }
function entriesBetween_(a, b) { return readRows_().map(rowToEntry_).filter(function (e) { return e.id && e.date >= a && e.date <= b; }); }

/* ---------- destinatários ---------- */
function gestorEmails_(cfg) { const a = (cfg.alerts || {}).gestor || ''; const list = a ? a.split(/[,;\s]+/).filter(String) : [];
  if (!list.length) { try { list.push(Session.getEffectiveUser().getEmail()); } catch (e) {} } return list.filter(String); }
function teamEmailsFor_(cfg, plat) { return (cfg.team || []).filter(function (t) { return t.email && (t.plats || []).indexOf(plat) >= 0; }).map(function (t) { return t.email; }); }

/* ---------- 1) alerta de horário vazio (roda a cada 15 min) ---------- */
function verificarHorariosVazios() {
  const cfg = readCfg_(); const al = cfg.alerts || {};
  if (!al.on) return;
  const now = new Date(); const iso = todayIso_();
  const hh = +Utilities.formatDate(now, tz_(), 'HH'), mm = +Utilities.formatDate(now, tz_(), 'mm');
  if (isNaN(hh) || isNaN(mm)) return; const mins = hh * 60 + mm;
  if (hh < (al.from == null ? 7 : al.from) || hh > (al.to == null ? 23 : al.to)) return;   // horário de silêncio
  const gp = gradePlats_(cfg); if (!gp.length) return;
  const pr = PropertiesService.getScriptProperties();
  const todays = entriesBetween_(iso, iso).filter(function (e) { return e.ativ === 'Postagem'; });
  slotsFor_(cfg, iso).forEach(function (slot) {
    const sm = +slot.slice(0, 2) * 60;
    if (mins < sm + 15 || mins > sm + 75) return;              // só avisa uma vez, logo depois dos 15 min
    gp.forEach(function (p) {
      const ok = todays.some(function (e) { return e.slot === slot && (e.plats || []).indexOf(p) >= 0; });
      const key = 'alert|' + iso + '|' + slot + '|' + p;
      if (ok || pr.getProperty(key)) return;
      const to = gestorEmails_(cfg).concat(al.team !== false ? teamEmailsFor_(cfg, p) : []);
      const uniq = to.filter(function (x, i) { return x && to.indexOf(x) === i; });
      if (!uniq.length) return;
      MailApp.sendEmail({to: uniq.join(','), subject: '⚠ CNG: ' + slot + ' ainda sem post no ' + p,
        htmlBody: mailWrap_('Horário sem post', '<p style="font-size:16px">O horário de <b>' + slot + '</b> de hoje (' + toBr_(iso) + ') passou de 15 minutos e ainda <b>não tem post registrado no ' + p + '</b>.</p>' +
          '<p>Se o post já saiu, registre no sistema. Se não, publique ou agende o quanto antes.</p>' + btn_('Abrir a Grade do dia', SITE_URL + '#grade'))});
      pr.setProperty(key, nowStr_()); pr.setProperty('last_alert', nowStr_() + ' · ' + slot + ' ' + p);
    });
  });
  // limpa marcações de dias anteriores
  const keys = pr.getKeys(); keys.forEach(function (k) { if (k.indexOf('alert|') === 0 && k.indexOf('alert|' + iso) !== 0) pr.deleteProperty(k); });
}
function sendTestAlert(name, token) {
  const u = must_(name, token); if (!u.admin) throw new Error('Só o gestor.');
  const cfg = readCfg_(); const to = gestorEmails_(cfg);
  if (!to.length) throw new Error('Cadastre o e-mail do gestor em Ajustes.');
  MailApp.sendEmail({to: to.join(','), subject: '✅ CNG: teste de alerta', htmlBody: mailWrap_('Teste de alerta', '<p>Os alertas de horário vazio vão chegar assim.</p>' + btn_('Abrir o sistema', SITE_URL))});
  return {ok: true, to: to.join(', ')};
}

/* ---------- 2) relatório semanal (segunda 7h30) ---------- */
function weekData_(a, b) {
  const cfg = readCfg_(); const all = entriesBetween_(addDaysIso_(a, -1), addDaysIso_(b, 30));
  const inWeek = all.filter(function (e) { return e.date >= a && e.date <= b; });
  let num = 0, den = 0; const late = {}; const met = {}; const gp = gradePlats_(cfg), mp = metaPlats_(cfg);
  gp.concat(mp).forEach(function (p) { met[p] = {n: 0, q: 0}; });
  for (let d = a; d <= b; d = addDaysIso_(d, 1)) {
    const day = inWeek.filter(function (e) { return e.date === d && e.ativ === 'Postagem'; });
    slotsFor_(cfg, d).forEach(function (sl) { gp.forEach(function (p) {
      den++; met[p].q++;
      if (day.some(function (e) { return e.slot === sl && (e.plats || []).indexOf(p) >= 0; })) { num++; met[p].n++; } else late[sl] = (late[sl] || 0) + 1; }); });
    mp.forEach(function (p) { const q = quota_(cfg, p, d), n = day.filter(function (e) { return (e.plats || []).indexOf(p) >= 0; }).length; num += Math.min(n, q); den += q; met[p].n += n; met[p].q += q; });
  }
  const reg = function (e) { return String(e.regAt || '').slice(0, 10); };
  const made = all.filter(function (e) { return reg(e) >= a && reg(e) <= b; });
  const people = (cfg.team || []).map(function (t) {
    const ns = made.filter(function (e) { return e.autor === t.name && e.ativ === 'Postagem'; });
    const vids = made.filter(function (e) { return e.autor === t.name && e.ativ === 'Vídeo (matéria)'; }).length;
    return {name: t.name, n: ns.length, vids: vids, vmeta: (+t.metaVideos || 0) * 7,
            link: ns.length ? Math.round(ns.filter(function (e) { return /^https?:/.test(e.link); }).length / ns.length * 100) : null};
  }).sort(function (x, y) { return y.n - x.n; });
  const acts = actRows_(); const today = todayIso_();
  return {cfg: cfg, pct: den ? Math.round(num / den * 100) : null, num: num, den: den, late: late, met: met, people: people,
          actsLate: acts.filter(function (x) { return x.status !== 'Concluído' && x.prazo && x.prazo < today; }),
          actsOpen: acts.filter(function (x) { return x.status !== 'Concluído'; })};
}
function enviarRelatorioSemanal() {
  const cfg = readCfg_(); if ((cfg.alerts || {}).weekly === false) return;
  const t = todayIso_(); const lastMon = addDaysIso_(t, -(((dowIso_(t) + 6) % 7) + 7));
  sendReport_(lastMon, addDaysIso_(lastMon, 6));
}
function sendTestReport(name, token) {
  const u = must_(name, token); if (!u.admin) throw new Error('Só o gestor.');
  const t = todayIso_(); const mon = addDaysIso_(t, -((dowIso_(t) + 6) % 7));
  return sendReport_(mon, t, true);
}
function sendReport_(a, b, parcial) {
  const w = weekData_(a, b); const to = gestorEmails_(w.cfg);
  if (!to.length) throw new Error('Cadastre o e-mail do gestor em Ajustes.');
  const pct = w.pct == null ? '—' : w.pct + '%';
  const color = w.pct >= 90 ? '#1f8a4c' : w.pct >= 70 ? '#b26a00' : '#c8102e';
  const lateTop = Object.keys(w.late).sort(function (x, y) { return w.late[y] - w.late[x]; }).slice(0, 3);
  const metRows = Object.keys(w.met).map(function (p) { const m = w.met[p]; const r = m.q ? Math.round(Math.min(m.n, m.q) / m.q * 100) : 0;
    return '<tr><td style="padding:6px 0">' + p + '</td><td style="text-align:right"><b>' + m.n + '</b> / ' + m.q + '</td><td style="text-align:right;color:' + (r >= 90 ? '#1f8a4c' : r >= 70 ? '#b26a00' : '#c8102e') + '"><b>' + r + '%</b></td></tr>'; }).join('');
  const pplRows = w.people.map(function (x, i) { return '<tr><td style="padding:6px 0">' + (i === 0 && x.n ? '🏆 ' : '') + x.name + '</td><td style="text-align:right"><b>' + x.n + '</b> notícias</td><td style="text-align:right">' + (x.link == null ? '—' : x.link + '% c/ link') + '</td><td style="text-align:right">' + (x.vmeta ? '🎥 ' + x.vids + '/' + x.vmeta : (x.vids ? '🎥 ' + x.vids : '')) + '</td></tr>'; }).join('');
  const pontos = [];
  if (w.pct != null && w.pct < 90) pontos.push('Cumprimento em ' + pct + ' (meta 90%).');
  if (lateTop.length) pontos.push('Horários que mais ficaram sem post: ' + lateTop.map(function (s) { return s + ' (' + w.late[s] + 'x)'; }).join(', ') + '.');
  Object.keys(w.met).forEach(function (p) { const m = w.met[p]; if (m.q && Math.min(m.n, m.q) / m.q < .9) pontos.push(p + ' abaixo da meta: ' + m.n + ' de ' + m.q + '.'); });
  w.people.filter(function (x) { return x.vmeta && x.vids < x.vmeta; }).forEach(function (x) { pontos.push(x.name + ': ' + x.vids + ' de ' + x.vmeta + ' vídeos.'); });
  const acts = w.actsLate.map(function (x) { return '<li><b>' + x.titulo + '</b> · ' + (x.resp || '—') + ' · prazo ' + toBr_(x.prazo) + '</li>'; }).join('');
  let met = ''; try { met = metricsSummary_(a, b); } catch (e) {}
  const body =
    '<p style="color:#666;margin:0">Semana de ' + toBr_(a) + ' a ' + toBr_(b) + (parcial ? ' (parcial, até hoje)' : '') + '</p>' +
    '<div style="background:#0d0d0d;color:#fff;border-radius:16px;padding:20px;margin:14px 0"><div style="color:#bbb;font-size:13px">Índice de cumprimento</div>' +
    '<div style="font-size:44px;font-weight:800;color:#c6e43c">' + pct + '</div><div style="font-size:14px">' + w.num + ' de ' + w.den + ' entregas esperadas</div></div>' +
    '<h3>Redes</h3><table style="width:100%;border-collapse:collapse;font-size:14px">' + metRows + '</table>' +
    '<h3>Produção por pessoa</h3><table style="width:100%;border-collapse:collapse;font-size:14px">' + pplRows + '</table>' +
    (met ? '<h3>Instagram (alcance)</h3>' + met : '') +
    '<h3>Pontos de atenção</h3>' + (pontos.length ? '<ul>' + pontos.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ul>' : '<p>Nenhum. A equipe bateu as metas. 👏</p>') +
    '<h3>Plano de ação</h3><p>' + w.actsOpen.length + ' ações abertas · <b style="color:' + (w.actsLate.length ? '#c8102e' : '#1f8a4c') + '">' + w.actsLate.length + ' atrasadas</b></p>' + (acts ? '<ul>' + acts + '</ul>' : '') +
    btn_('Abrir o Desempenho', SITE_URL + '#desempenho');
  MailApp.sendEmail({to: to.join(','), subject: '📊 CNG: relatório da semana ' + toBr_(a).slice(0, 5) + ' a ' + toBr_(b).slice(0, 5) + ' · ' + pct, htmlBody: mailWrap_('Relatório semanal', body)});
  PropertiesService.getScriptProperties().setProperty('last_report', nowStr_());
  return {ok: true, to: to.join(', ')};
}
function mailWrap_(title, inner) { return '<div style="font-family:Arial,sans-serif;max-width:620px;margin:auto;color:#14171c"><div style="border-bottom:3px solid #c6e43c;padding-bottom:8px;margin-bottom:12px"><b style="font-size:18px">CNG Notícias</b> · <span style="color:#666">' + title + '</span></div>' + inner + '<p style="color:#999;font-size:12px;margin-top:24px">Mensagem automática do Controle de Postagens.</p></div>'; }
function btn_(t, u) { return '<p style="margin:18px 0"><a href="' + u + '" style="background:#c6e43c;color:#0d0d0d;text-decoration:none;font-weight:700;padding:10px 16px;border-radius:10px">' + t + '</a></p>'; }

/* ---------- 3) métricas reais do Instagram (Windsor.ai) ---------- */
function setWindsorKey(name, token, key) {
  const u = must_(name, token); if (!u.admin) throw new Error('Só o gestor.');
  key = String(key || '').trim();
  const pr = PropertiesService.getScriptProperties();
  if (!key) { pr.deleteProperty('WINDSOR_KEY'); return {ok: true, removed: true}; }
  if (!/^[A-Za-z0-9_\-]{16,200}$/.test(key)) throw new Error('Chave com formato inválido.');
  pr.setProperty('WINDSOR_KEY', key);
  return refreshMetrics(name, token);
}
function refreshMetrics(name, token) { const u = must_(name, token); if (!u.admin) throw new Error('Só o gestor.'); return atualizarMetricas(); }
function atualizarMetricas() {
  const pr = PropertiesService.getScriptProperties(); const key = pr.getProperty('WINDSOR_KEY');
  if (!key) return {ok: false, msg: 'Sem chave do Windsor.ai.'};
  const fields = ['date', 'timestamp', 'media_id', 'media_type', 'media_product_type', 'media_caption', 'media_permalink', 'media_reach', 'media_views', 'media_engagement', 'media_shares', 'media_saved', 'media_total_like_count', 'media_total_comments_count'];
  const url = 'https://connectors.windsor.ai/instagram?api_key=' + encodeURIComponent(key) + '&date_preset=last_60d&fields=' + fields.join(',');
  try {
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (res.getResponseCode() !== 200) throw new Error('Windsor respondeu ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 160));
    const js = JSON.parse(res.getContentText()); const rows = Array.isArray(js) ? js : (js.data || []);
    const seen = {}; const out = [];
    rows.forEach(function (r) { const id = String(r.media_id || r.media_permalink || ''); if (!id || seen[id]) return; seen[id] = 1;
      out.push([id, String(r.timestamp || r.date || ''), String(r.media_type || ''), String(r.media_product_type || ''), String(r.media_caption || '').slice(0, 300),
                String(r.media_permalink || ''), +r.media_reach || 0, +r.media_views || 0, +r.media_engagement || 0, +r.media_shares || 0, +r.media_saved || 0,
                +r.media_total_like_count || 0, +r.media_total_comments_count || 0]); });
    const ss = ss_(); let sh = ss.getSheetByName(ABA_MET);
    if (!sh) { sh = ss.insertSheet(ABA_MET); }
    sh.clearContents();
    const head = ['ID', 'Publicado em', 'Tipo', 'Formato', 'Legenda', 'Link', 'Alcance', 'Visualizações', 'Engajamento', 'Compartilhamentos', 'Salvos', 'Curtidas', 'Comentários'];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    if (out.length) { const rg = sh.getRange(2, 1, out.length, head.length); rg.setNumberFormat('@'); rg.setValues(out.map(function (r) { return r.map(String); })); }
    pr.setProperty('metrics_at', nowStr_()); pr.deleteProperty('metrics_err');
    return {ok: true, posts: out.length, at: nowStr_()};
  } catch (e) {
    pr.setProperty('metrics_err', nowStr_() + ' · ' + String(e.message || e).slice(0, 200));
    throw e;
  }
}
function getMetrics(name, token) {
  must_(name, token);
  const sh = ss_().getSheetByName(ABA_MET); const pr = PropertiesService.getScriptProperties();
  const meta = {at: pr.getProperty('metrics_at') || '', err: pr.getProperty('metrics_err') || '', hasKey: !!pr.getProperty('WINDSOR_KEY')};
  if (!sh || sh.getLastRow() < 2) return {meta: meta, posts: []};
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
  return {meta: meta, posts: v.map(function (r) { return {id: String(r[0]), ts: String(r[1]), type: String(r[2]), product: String(r[3]), caption: String(r[4]), link: String(r[5]),
    reach: +r[6] || 0, views: +r[7] || 0, eng: +r[8] || 0, shares: +r[9] || 0, saved: +r[10] || 0, likes: +r[11] || 0, comments: +r[12] || 0}; })};
}
function metricsSummary_(a, b) {
  const sh = ss_().getSheetByName(ABA_MET); if (!sh || sh.getLastRow() < 2) return '';
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues().filter(function (r) { const d = String(r[1]).slice(0, 10); return d >= a && d <= b; });
  if (!v.length) return '';
  const reach = v.reduce(function (x, r) { return x + (+r[6] || 0); }, 0), eng = v.reduce(function (x, r) { return x + (+r[8] || 0); }, 0);
  const top = v.sort(function (x, y) { return (+y[6] || 0) - (+x[6] || 0); }).slice(0, 3);
  return '<p>' + v.length + ' posts · alcance somado <b>' + reach.toLocaleString('pt-BR') + '</b> · engajamento <b>' + eng.toLocaleString('pt-BR') + '</b></p><ol>' +
    top.map(function (r) { return '<li><a href="' + r[5] + '">' + String(r[4] || 'post').slice(0, 70) + '</a> · ' + (+r[6] || 0).toLocaleString('pt-BR') + ' de alcance</li>'; }).join('') + '</ol>';
}
