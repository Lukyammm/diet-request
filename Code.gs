/*************************************************************************
 * DIET REQUEST — Solicitação de Dieta (Nutrição)
 * Google Apps Script — servidor
 *
 * INSTALAÇÃO
 *  1. Abra a planilha Google que servirá de banco de dados.
 *  2. Extensões > Apps Script.
 *  3. Cole este arquivo como "Code.gs" e crie um arquivo HTML "index"
 *     com o conteúdo de index.html.
 *  4. Execute a função setup() uma vez (autorize os acessos).
 *  5. Implantar > Nova implantação > App da Web
 *       Executar como: Eu
 *       Quem tem acesso: Qualquer pessoa (o app tem login próprio).
 *  6. Primeiro acesso administrativo: login "admin", senha "dieta@2026"
 *     (troca obrigatória no primeiro login).
 *
 * PAINEL DE TV: abra a URL do app com ?view=tv  (não exige login).
 *************************************************************************/

var SH = { SOL: 'SOLICITACOES', USU: 'USUARIOS', LIS: 'LISTAS' };

var SOL_HEADERS = [
  'ID', 'CRIADO_EM', 'STATUS', 'TIPO',
  'PACIENTE', 'PRONTUARIO', 'NASCIMENTO', 'CLINICA', 'ENFERMARIA_LEITO', 'ISOLAMENTO',
  'PROFISSIONAL', 'CATEGORIA',
  'HORARIO_PRESCRICAO', 'TIPO_DIETA', 'PRESCRICAO', 'COMORBIDADES', 'QTD_ENTERAL',
  'ESPESSADA', 'CONSISTENCIA', 'ESPESSANTE_AGUA',
  'MOTIVO_SUSPENSAO', 'NOVA_PRESCRICAO',
  'MAMADEIRA', 'ESPESSAMENTO',
  'VIA_ORAL', 'SUCO_MANITOL',
  'SOLIC_FONO', 'OBS',
  'ANALISE', 'MOTIVO_RECUSA', 'JUSTIFICATIVA', 'DIETA_ENTREGUE',
  'RESPONSAVEL', 'RESOLVIDO_EM', 'TEMPO_MIN', 'CONFORME'
];

var USU_HEADERS = ['ID', 'NOME', 'LOGIN', 'HASH', 'SALT', 'PERFIL', 'ATIVO', 'PRIMEIRO_ACESSO', 'TOKEN', 'TOKEN_EXP'];

var TIPOS = [
  'LIBERAÇÃO DE DIETA',
  'SUSPENSÃO DE DIETA',
  'MUDANÇA DE PRESCRIÇÃO',
  'DIETA TESTE-FONO',
  'PREPARO PARA EXAMES'
];

/* ======================= WEB APP ======================= */

function doGet(e) {
  ensureSetup_();
  var t = HtmlService.createTemplateFromFile('index');
  t.mode = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : '';
  return t.evaluate()
    .setTitle('Solicitação de Dieta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ======================= SETUP ======================= */

function setup() {
  PropertiesService.getScriptProperties().deleteProperty('SETUP_OK');
  ensureSetup_();
  return 'Setup concluído. Login inicial: admin / dieta@2026';
}

function ensureSetup_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SETUP_OK') === '1') return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (props.getProperty('SETUP_OK') === '1') return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sol = getOrCreate_(ss, SH.SOL, SOL_HEADERS);
    var usu = getOrCreate_(ss, SH.USU, USU_HEADERS);
    var lis = getOrCreate_(ss, SH.LIS, ['TIPO', 'VALOR']);

    if (lis.getLastRow() < 2) seedListas_(lis);
    if (usu.getLastRow() < 2) seedAdmin_(usu);

    sol.setFrozenRows(1); usu.setFrozenRows(1); lis.setFrozenRows(1);
    props.setProperty('SETUP_OK', '1');
  } finally {
    lock.releaseLock();
  }
}

function getOrCreate_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var first = sh.getRange(1, 1).getValue();
  if (String(first) !== headers[0]) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}

function seedListas_(sh) {
  var seed = [];
  var push = function (tipo, arr) { arr.forEach(function (v) { seed.push([tipo, v]); }); };
  push('CLINICA', ['UTI ADULTO', 'UTI PEDIÁTRICA', 'UTI NEONATAL', 'CLÍNICA MÉDICA', 'CLÍNICA CIRÚRGICA', 'PEDIATRIA', 'CARDIOLOGIA PEDIÁTRICA', 'OBSERVAÇÃO', 'EMERGÊNCIA']);
  push('CATEGORIA', ['MÉDICO(A)', 'ENFERMEIRO(A)', 'TÉC. DE ENFERMAGEM', 'NUTRICIONISTA', 'FISIOTERAPEUTA', 'FONOAUDIÓLOGO(A)', 'OUTROS']);
  push('MOTIVO_SUSPENSAO', ['VÔMITO', 'PERDA DE SONDA', 'EXAME', 'CIRURGIA', 'INSTABILIDADE CLÍNICA', 'ALTA', 'MUDANÇA DE VIA']);
  push('MAMADEIRA', ['MAM', 'BICO X', 'CHUQUINHA', 'BICO LONGO', 'BICO ORTODÔNTICO', 'OUTROS']);
  push('CONSISTENCIA', ['NÉCTAR', 'MEL', 'PUDIM']);
  push('MOTIVO_RECUSA', [
    'ENTREGA NO PRÓXIMO HORÁRIO PADRÃO DE DISTRIBUIÇÃO',
    'HORÁRIO DE PENDÊNCIA EXPIRADO',
    'PRESCRIÇÃO MÉDICA DE DIETA ZERO',
    'SEM PRESCRIÇÃO MÉDICA DE DIETA',
    'SOLICITAÇÃO DUPLICADA',
    'VIA DE ALIMENTAÇÃO DIFERENTE DA PRESCRITA'
  ]);
  seed.push(['CONFIG', 'SLA_MINUTOS=30']);
  sh.getRange(2, 1, seed.length, 2).setValues(seed);
}

function seedAdmin_(sh) {
  var salt = Utilities.getUuid();
  sh.appendRow(['U1', 'ADMINISTRADOR', 'admin', hash_('dieta@2026', salt), salt, 'ADMIN', 'SIM', 'SIM', '', '']);
}

/* ======================= UTILS ======================= */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { ensureSetup_(); return ss_().getSheetByName(name); }
function nowIso_() { return new Date().toISOString(); }
function s_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function rowsAsObjects_(sh, headers) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, headers.length).getValues();
  return data.map(function (row, i) {
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = s_(row[j]); });
    return o;
  });
}

function listas_() {
  var sh = sheet_(SH.LIS);
  var out = {};
  var last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    var tipo = s_(r[0]), val = s_(r[1]);
    if (!tipo || !val) return;
    if (!out[tipo]) out[tipo] = [];
    out[tipo].push(val);
  });
  return out;
}

function slaMin_() {
  var cfg = (listas_().CONFIG || []);
  for (var i = 0; i < cfg.length; i++) {
    var m = cfg[i].match(/^SLA_MINUTOS\s*=\s*(\d+)/i);
    if (m) return parseInt(m[1], 10);
  }
  return 30;
}

function nextId_() {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty('SEQ') || '0', 10) + 1;
  props.setProperty('SEQ', String(n));
  return 'SD-' + ('0000' + n).slice(-4);
}

function hash_(senha, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + senha, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

/* ======================= AUTENTICAÇÃO ======================= */

function login(loginName, senha) {
  loginName = s_(loginName).toLowerCase(); senha = s_(senha);
  if (!loginName || !senha) return { ok: false, erro: 'Informe login e senha.' };
  var sh = sheet_(SH.USU);
  var users = rowsAsObjects_(sh, USU_HEADERS);
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (u.LOGIN.toLowerCase() !== loginName) continue;
    if (u.ATIVO !== 'SIM') return { ok: false, erro: 'Usuário inativo. Procure o administrador.' };
    if (hash_(senha, u.SALT) !== u.HASH) return { ok: false, erro: 'Login ou senha incorretos.' };
    var token = Utilities.getUuid();
    var exp = Date.now() + 12 * 3600 * 1000;
    sh.getRange(u._row, 9, 1, 2).setValues([[token, String(exp)]]);
    return { ok: true, token: token, nome: u.NOME, perfil: u.PERFIL, primeiroAcesso: u.PRIMEIRO_ACESSO === 'SIM' };
  }
  return { ok: false, erro: 'Login ou senha incorretos.' };
}

function auth_(token) {
  token = s_(token);
  if (!token) return null;
  var users = rowsAsObjects_(sheet_(SH.USU), USU_HEADERS);
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (u.TOKEN === token && u.ATIVO === 'SIM' && Number(u.TOKEN_EXP) > Date.now()) return u;
  }
  return null;
}

function apiMe(token) {
  var u = auth_(token);
  if (!u) return { ok: false };
  return { ok: true, nome: u.NOME, perfil: u.PERFIL, primeiroAcesso: u.PRIMEIRO_ACESSO === 'SIM' };
}

function logout(token) {
  var u = auth_(token);
  if (u) sheet_(SH.USU).getRange(u._row, 9, 1, 2).setValues([['', '']]);
  return { ok: true };
}

function trocarSenha(token, nova) {
  var u = auth_(token);
  if (!u) return { ok: false, erro: 'Sessão expirada. Entre novamente.' };
  nova = s_(nova);
  if (nova.length < 6) return { ok: false, erro: 'A nova senha precisa de pelo menos 6 caracteres.' };
  var salt = Utilities.getUuid();
  sheet_(SH.USU).getRange(u._row, 4, 1, 2).setValues([[hash_(nova, salt), salt]]);
  sheet_(SH.USU).getRange(u._row, 8).setValue('NÃO');
  return { ok: true };
}

/* ======================= PÚBLICO ======================= */

function pubBoot() {
  var L = listas_();
  return {
    ok: true,
    tipos: TIPOS,
    clinicas: L.CLINICA || [],
    categorias: L.CATEGORIA || [],
    motivosSuspensao: L.MOTIVO_SUSPENSAO || [],
    mamadeiras: L.MAMADEIRA || [],
    consistencias: L.CONSISTENCIA || [],
    slaMin: slaMin_()
  };
}

function pubCriar(o) {
  o = o || {};
  var req = function (campo, rotulo) {
    if (!s_(o[campo])) throw new Error('Campo obrigatório: ' + rotulo + '.');
  };
  try {
    req('tipo', 'tipo de solicitação');
    if (TIPOS.indexOf(o.tipo) === -1) throw new Error('Tipo de solicitação inválido.');
    req('paciente', 'nome do paciente');
    req('prontuario', 'prontuário');
    req('clinica', 'clínica');
    req('enfermariaLeito', 'enfermaria e leito');
    req('profissional', 'profissional solicitante');
    req('categoria', 'categoria profissional');

    if (o.tipo === TIPOS[0]) { req('tipoDieta', 'tipo de dieta'); req('prescricao', 'prescrição médica'); }
    if (o.tipo === TIPOS[1]) { req('motivoSuspensao', 'motivo da suspensão'); }
    if (o.tipo === TIPOS[2]) { req('novaPrescricao', 'nova prescrição'); }
    if (o.tipo === TIPOS[3]) { req('mamadeira', 'mamadeira / bico'); }
    if (o.tipo === TIPOS[4]) { req('viaOral', 'via oral'); }

    var sucoManitol = '';
    if (o.tipo === TIPOS[4]) {
      sucoManitol = (s_(o.viaOral) === 'SIM') ? 'SIM — SUCO DE LARANJA (500 ML)' : 'NÃO SE APLICA';
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var id = nextId_();
      var row = [
        id, nowIso_(), 'PENDENTE', s_(o.tipo),
        s_(o.paciente).toUpperCase(), s_(o.prontuario), s_(o.nascimento), s_(o.clinica), s_(o.enfermariaLeito).toUpperCase(), s_(o.isolamento),
        s_(o.profissional).toUpperCase(), s_(o.categoria),
        s_(o.horarioPrescricao), s_(o.tipoDieta), s_(o.prescricao), s_(o.comorbidades), s_(o.qtdEnteral),
        s_(o.espessada), s_(o.consistencia), s_(o.espessanteAgua),
        s_(o.motivoSuspensao), s_(o.novaPrescricao),
        s_(o.mamadeira), s_(o.espessamento),
        s_(o.viaOral), sucoManitol,
        s_(o.solicFono), s_(o.obs),
        '', '', '', '', '', '', '', ''
      ];
      sheet_(SH.SOL).appendRow(row);
      return { ok: true, id: id };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { ok: false, erro: e.message || String(e) };
  }
}

function pubPainel() {
  var all = rowsAsObjects_(sheet_(SH.SOL), SOL_HEADERS);
  var hoje = new Date().toISOString().slice(0, 10);
  var pend = [], recentes = [], stats = { total: 0, pendentes: 0, atendidas: 0, recusadas: 0 };

  all.forEach(function (r) {
    if (r.CRIADO_EM.slice(0, 10) === hoje) {
      stats.total++;
      if (r.STATUS === 'PENDENTE') stats.pendentes++;
      if (r.STATUS === 'ATENDIDA') stats.atendidas++;
      if (r.STATUS === 'RECUSADA') stats.recusadas++;
    }
    var resumo = r.PRESCRICAO || r.NOVA_PRESCRICAO || r.MOTIVO_SUSPENSAO || r.MAMADEIRA || r.SUCO_MANITOL || '';
    if (r.STATUS === 'PENDENTE') {
      pend.push({ id: r.ID, criadoEm: r.CRIADO_EM, tipo: r.TIPO, paciente: r.PACIENTE, clinica: r.CLINICA, leito: r.ENFERMARIA_LEITO, resumo: resumo });
    } else {
      recentes.push({ id: r.ID, tipo: r.TIPO, paciente: r.PACIENTE, status: r.STATUS, resolvidoEm: r.RESOLVIDO_EM, tempoMin: r.TEMPO_MIN });
    }
  });

  pend.sort(function (a, b) { return a.criadoEm < b.criadoEm ? -1 : 1; });
  recentes.sort(function (a, b) { return a.resolvidoEm > b.resolvidoEm ? -1 : 1; });
  return { ok: true, pendentes: pend, recentes: recentes.slice(0, 8), stats: stats, slaMin: slaMin_(), agora: nowIso_() };
}

/* ======================= OPERAÇÃO (LOGIN) ======================= */

function apiPendentes(token) {
  if (!auth_(token)) return { ok: false, erro: 'Sessão expirada.' };
  var pend = rowsAsObjects_(sheet_(SH.SOL), SOL_HEADERS)
    .filter(function (r) { return r.STATUS === 'PENDENTE'; })
    .sort(function (a, b) { return a.CRIADO_EM < b.CRIADO_EM ? -1 : 1; });
  return { ok: true, itens: pend, slaMin: slaMin_(), motivosRecusa: listas_().MOTIVO_RECUSA || [] };
}

function apiResolver(token, o) {
  var u = auth_(token);
  if (!u) return { ok: false, erro: 'Sessão expirada.' };
  o = o || {};
  var analise = s_(o.analise);
  if (analise !== 'ACEITA' && analise !== 'RECUSADA') return { ok: false, erro: 'Análise inválida.' };
  if (analise === 'RECUSADA' && !s_(o.motivoRecusa)) return { ok: false, erro: 'Informe o motivo padronizado da recusa.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SH.SOL);
    var all = rowsAsObjects_(sh, SOL_HEADERS);
    for (var i = 0; i < all.length; i++) {
      var r = all[i];
      if (r.ID !== s_(o.id)) continue;
      if (r.STATUS !== 'PENDENTE') return { ok: false, erro: 'Esta solicitação já foi resolvida por outra pessoa.' };
      var agora = new Date();
      var tempoMin = Math.max(0, Math.round((agora.getTime() - new Date(r.CRIADO_EM).getTime()) / 60000));
      var conforme = analise === 'ACEITA' ? (tempoMin <= slaMin_() ? 'SIM' : 'NÃO') : '';
      var status = analise === 'ACEITA' ? 'ATENDIDA' : 'RECUSADA';
      var col = function (h) { return SOL_HEADERS.indexOf(h) + 1; };
      sh.getRange(r._row, col('STATUS')).setValue(status);
      sh.getRange(r._row, col('ANALISE'), 1, 8).setValues([[
        analise,
        analise === 'RECUSADA' ? s_(o.motivoRecusa) : '',
        s_(o.justificativa),
        analise === 'ACEITA' ? s_(o.dietaEntregue) : '',
        u.NOME,
        agora.toISOString(),
        String(tempoMin),
        conforme
      ]]);
      return { ok: true };
    }
    return { ok: false, erro: 'Solicitação não encontrada.' };
  } finally {
    lock.releaseLock();
  }
}

function apiHistorico(token, f) {
  if (!auth_(token)) return { ok: false, erro: 'Sessão expirada.' };
  f = f || {};
  var mes = s_(f.mes), status = s_(f.status), busca = s_(f.busca).toLowerCase();
  var itens = rowsAsObjects_(sheet_(SH.SOL), SOL_HEADERS)
    .filter(function (r) {
      if (mes && r.CRIADO_EM.slice(0, 7) !== mes) return false;
      if (status && r.STATUS !== status) return false;
      if (busca) {
        var alvo = (r.PACIENTE + ' ' + r.PRONTUARIO + ' ' + r.ID).toLowerCase();
        if (alvo.indexOf(busca) === -1) return false;
      }
      return true;
    })
    .sort(function (a, b) { return a.CRIADO_EM > b.CRIADO_EM ? -1 : 1; })
    .slice(0, 300);
  return { ok: true, itens: itens };
}

function apiRelatorio(token, mes) {
  if (!auth_(token)) return { ok: false, erro: 'Sessão expirada.' };
  mes = s_(mes);
  var itens = rowsAsObjects_(sheet_(SH.SOL), SOL_HEADERS)
    .filter(function (r) { return !mes || r.CRIADO_EM.slice(0, 7) === mes; });

  var rel = {
    ok: true, mes: mes, total: itens.length,
    pendentes: 0, atendidas: 0, recusadas: 0,
    tempoMedioMin: 0, conformePct: 0,
    porTipo: {}, porClinica: {}, recusaPorMotivo: {}, porCategoria: {}
  };
  var tempos = [], conf = 0, confBase = 0;
  itens.forEach(function (r) {
    if (r.STATUS === 'PENDENTE') rel.pendentes++;
    if (r.STATUS === 'ATENDIDA') rel.atendidas++;
    if (r.STATUS === 'RECUSADA') rel.recusadas++;
    rel.porTipo[r.TIPO] = (rel.porTipo[r.TIPO] || 0) + 1;
    if (r.CLINICA) rel.porClinica[r.CLINICA] = (rel.porClinica[r.CLINICA] || 0) + 1;
    if (r.CATEGORIA) rel.porCategoria[r.CATEGORIA] = (rel.porCategoria[r.CATEGORIA] || 0) + 1;
    if (r.STATUS === 'RECUSADA' && r.MOTIVO_RECUSA) rel.recusaPorMotivo[r.MOTIVO_RECUSA] = (rel.recusaPorMotivo[r.MOTIVO_RECUSA] || 0) + 1;
    if (r.STATUS === 'ATENDIDA' && r.TEMPO_MIN !== '') {
      tempos.push(Number(r.TEMPO_MIN)); confBase++;
      if (r.CONFORME === 'SIM') conf++;
    }
  });
  if (tempos.length) rel.tempoMedioMin = Math.round(tempos.reduce(function (a, b) { return a + b; }, 0) / tempos.length);
  if (confBase) rel.conformePct = Math.round(100 * conf / confBase);
  rel.slaMin = slaMin_();
  return rel;
}

/* ======================= ADMIN ======================= */

function requireAdmin_(token) {
  var u = auth_(token);
  if (!u || u.PERFIL !== 'ADMIN') return null;
  return u;
}

function apiUsuarios(token) {
  if (!requireAdmin_(token)) return { ok: false, erro: 'Apenas administradores.' };
  var itens = rowsAsObjects_(sheet_(SH.USU), USU_HEADERS).map(function (u) {
    return { id: u.ID, nome: u.NOME, login: u.LOGIN, perfil: u.PERFIL, ativo: u.ATIVO === 'SIM', primeiroAcesso: u.PRIMEIRO_ACESSO === 'SIM' };
  });
  return { ok: true, itens: itens };
}

function apiSalvarUsuario(token, o) {
  var adm = requireAdmin_(token);
  if (!adm) return { ok: false, erro: 'Apenas administradores.' };
  o = o || {};
  var nome = s_(o.nome).toUpperCase(), loginName = s_(o.login).toLowerCase(), perfil = s_(o.perfil) === 'ADMIN' ? 'ADMIN' : 'NUTRICAO';
  if (!nome || !loginName) return { ok: false, erro: 'Nome e login são obrigatórios.' };
  var sh = sheet_(SH.USU);
  var users = rowsAsObjects_(sh, USU_HEADERS);

  if (o.id) {
    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      if (u.ID !== s_(o.id)) continue;
      if (u.ID === adm.ID && o.ativo === false) return { ok: false, erro: 'Você não pode desativar o próprio usuário.' };
      sh.getRange(u._row, 2).setValue(nome);
      sh.getRange(u._row, 6, 1, 2).setValues([[perfil, o.ativo === false ? 'NÃO' : 'SIM']]);
      if (s_(o.senha)) {
        var salt = Utilities.getUuid();
        sh.getRange(u._row, 4, 1, 2).setValues([[hash_(s_(o.senha), salt), salt]]);
        sh.getRange(u._row, 8).setValue('SIM');
        sh.getRange(u._row, 9, 1, 2).setValues([['', '']]);
      }
      return { ok: true };
    }
    return { ok: false, erro: 'Usuário não encontrado.' };
  }

  for (var j = 0; j < users.length; j++) {
    if (users[j].LOGIN.toLowerCase() === loginName) return { ok: false, erro: 'Este login já existe.' };
  }
  var senha = s_(o.senha) || 'dieta@2026';
  var salt2 = Utilities.getUuid();
  var id = 'U' + (users.length + 1) + '-' + Date.now().toString(36);
  sh.appendRow([id, nome, loginName, hash_(senha, salt2), salt2, perfil, 'SIM', 'SIM', '', '']);
  return { ok: true };
}

var LISTAS_EDITAVEIS = ['CLINICA', 'CATEGORIA', 'MOTIVO_SUSPENSAO', 'MAMADEIRA', 'CONSISTENCIA', 'MOTIVO_RECUSA'];

function apiListas(token) {
  if (!requireAdmin_(token)) return { ok: false, erro: 'Apenas administradores.' };
  var L = listas_();
  var out = { ok: true, listas: {}, slaMin: slaMin_() };
  LISTAS_EDITAVEIS.forEach(function (t) { out.listas[t] = L[t] || []; });
  return out;
}

function apiSalvarLista(token, tipo, valores) {
  if (!requireAdmin_(token)) return { ok: false, erro: 'Apenas administradores.' };
  tipo = s_(tipo);
  if (LISTAS_EDITAVEIS.indexOf(tipo) === -1) return { ok: false, erro: 'Lista inválida.' };
  valores = (valores || []).map(s_).filter(function (v) { return !!v; });
  if (!valores.length) return { ok: false, erro: 'A lista não pode ficar vazia.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SH.LIS);
    var last = sh.getLastRow();
    var keep = [];
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
        if (s_(r[0]) !== tipo && s_(r[0])) keep.push([s_(r[0]), s_(r[1])]);
      });
    }
    valores.forEach(function (v) { keep.push([tipo, v.toUpperCase()]); });
    if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
    sh.getRange(2, 1, keep.length, 2).setValues(keep);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiSalvarSla(token, minutos) {
  if (!requireAdmin_(token)) return { ok: false, erro: 'Apenas administradores.' };
  var n = parseInt(minutos, 10);
  if (!(n > 0 && n <= 720)) return { ok: false, erro: 'Informe um SLA entre 1 e 720 minutos.' };
  var sh = sheet_(SH.LIS);
  var last = sh.getLastRow();
  for (var i = 2; i <= last; i++) {
    if (s_(sh.getRange(i, 1).getValue()) === 'CONFIG' && /^SLA_MINUTOS/i.test(s_(sh.getRange(i, 2).getValue()))) {
      sh.getRange(i, 2).setValue('SLA_MINUTOS=' + n);
      return { ok: true };
    }
  }
  sh.appendRow(['CONFIG', 'SLA_MINUTOS=' + n]);
  return { ok: true };
}
