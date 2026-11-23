/**
 * Backend do convite da Vitoria (custo zero).
 * Google Sheets + Apps Script, publicado como Web App.
 *
 * Folhas usadas no Sheet "RSVP Vitoria":
 *   respostas  : id | timestamp | status | adultos | criancas | n_adultos | n_criancas | nome_contacto | nota | convidado_id | user_agent
 *   convidados : convidado_id | nome | status | adultos | criancas | nota | resposta_id | criado_em | atualizado_em
 *
 * Ações:
 *   doPost {action:"rsvp", ...}        -> pública (sem chave): grava a resposta e liga-a a um convidado.
 *   doPost {action:"list", key}        -> admin: devolve respostas e convidados.
 *   doPost {action:"addGuests", key}   -> admin: adiciona convidados "pendente".
 *   doPost {action:"updateGuest", key} -> admin: edita nome/estado/adultos/criancas/nota.
 *   doPost {action:"deleteGuest", key} -> admin: apaga um convidado.
 *   doPost {action:"linkResponse", key}-> admin: liga uma resposta a um convidado (ou cria convidado).
 *   doGet                              -> verificação simples: {ok:true}.
 *
 * A chave de admin fica em Definições do projeto -> Propriedades do script -> ADMIN_KEY.
 */

var SHEET_NAME = 'RSVP Vitoria';
var H_RESP = ['id','timestamp','status','adultos','criancas','n_adultos','n_criancas','nome_contacto','nota','convidado_id','user_agent'];
var H_CONV = ['convidado_id','nome','status','adultos','criancas','nota','resposta_id','criado_em','atualizado_em'];
var MAX_LEN = 4000;   // limite de segurança por campo de texto

/* ============================ entradas ============================ */

function doGet(e){
  return json({ ok:true, service:'RSVP Vitoria' });
}

function doPost(e){
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json({ ok:false, error:'json invalido' });
  }
  var action = String(data.action || '');

  // Ação pública (sem chave).
  if (action === 'rsvp') return handleRsvp(data);

  // Ações de admin (exigem chave correta).
  if (['list','addGuests','updateGuest','deleteGuest','linkResponse'].indexOf(action) >= 0) {
    if (!checkKey(data.key)) return json({ ok:false, error:'auth' });
    if (action === 'list')         return handleList();
    if (action === 'addGuests')    return withLock(function(){ return handleAddGuests(data); });
    if (action === 'updateGuest')  return withLock(function(){ return handleUpdateGuest(data); });
    if (action === 'deleteGuest')  return withLock(function(){ return handleDeleteGuest(data); });
    if (action === 'linkResponse') return withLock(function(){ return handleLinkResponse(data); });
  }
  return json({ ok:false, error:'accao desconhecida' });
}

/* ============================ RSVP (público) ============================ */

function handleRsvp(data){
  return withLock(function(){
    var status = (data.status === 'nao') ? 'nao' : (data.status === 'vai' ? 'vai' : null);
    if (!status) return json({ ok:false, error:'status em falta' });

    var adultos  = cleanNames(data.adultos);
    var criancas = cleanNames(data.criancas);
    var nome     = clip(String(data.nome_contacto || '').trim());
    var nota     = clip(String(data.nota || '').trim());

    if (status === 'vai' && adultos.length < 1) return json({ ok:false, error:'precisa de pelo menos 1 adulto' });
    if (status === 'nao' && !nome)              return json({ ok:false, error:'precisa do nome' });
    if (status === 'vai' && !nome) nome = adultos[0] || '';

    var id = clip(String(data.id || '').trim()) || Utilities.getUuid();
    var ua = clip(String(data.user_agent || ''));
    var now = new Date();

    var rs = sheet('respostas', H_RESP);
    var rows = rs.getDataRange().getValues();
    var rIdx = findRow(rows, 0, id);           // upsert por id

    var convId = rIdx > 0 ? String(rows[rIdx][9] || '') : '';
    var record = [ id, now, status, adultos.join(', '), criancas.join(', '),
                   adultos.length, criancas.length, nome, nota, convId, ua ];

    if (rIdx > 0) rs.getRange(rIdx+1, 1, 1, H_RESP.length).setValues([record]);
    else          rs.appendRow(record);

    // Ligar a um convidado (atualiza o existente ou cria um novo).
    linkRsvpToGuest(id, status, adultos, criancas, nota, nome, now);

    if (typeof notifyEmail === 'function') { try { notifyEmail(status, nome, adultos, criancas, nota); } catch (err) {} }
    return json({ ok:true });
  });
}

/**
 * Procura um convidado cujo nome bata certo (sem acentos/maiúsculas) com o
 * nome de contacto ou com um dos adultos; atualiza-o. Se não existir, cria.
 */
function linkRsvpToGuest(respId, status, adultos, criancas, nota, nome, now){
  var cs = sheet('convidados', H_CONV);
  var rows = cs.getDataRange().getValues();

  var keys = [nome].concat(adultos).map(norm).filter(String);
  var matchIdx = -1;
  for (var i = 1; i < rows.length; i++){
    var gn = norm(rows[i][1]);
    if (!gn) continue;
    if (keys.indexOf(gn) >= 0){ matchIdx = i; break; }
  }

  if (matchIdx > 0){
    var convId = String(rows[matchIdx][0]);
    cs.getRange(matchIdx+1, 2, 1, 8).setValues([[
      rows[matchIdx][1] || nome, status, adultos.join(', '), criancas.join(', '),
      nota, respId, rows[matchIdx][7] || now, now
    ]]);
    setRespConv(respId, convId);
  } else {
    var newId = Utilities.getUuid();
    cs.appendRow([ newId, nome || 'Sem nome', status, adultos.join(', '), criancas.join(', '), nota, respId, now, now ]);
    setRespConv(respId, newId);
  }
}

function setRespConv(respId, convId){
  var rs = sheet('respostas', H_RESP);
  var rows = rs.getDataRange().getValues();
  var i = findRow(rows, 0, respId);
  if (i > 0) rs.getRange(i+1, 10).setValue(convId);
}

/* ============================ admin ============================ */

function handleList(){
  var resp = readSheet('respostas', H_RESP);
  var conv = readSheet('convidados', H_CONV);
  return json({ ok:true, respostas:resp, convidados:conv });
}

function handleAddGuests(data){
  var names = (data.names || []).map(function(n){ return clip(String(n||'').trim()); }).filter(String);
  if (!names.length) return json({ ok:false, error:'sem nomes' });

  var cs = sheet('convidados', H_CONV);
  var rows = cs.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < rows.length; i++) existing[norm(rows[i][1])] = true;

  var now = new Date(), added = 0;
  names.forEach(function(nome){
    if (existing[norm(nome)]) return;
    existing[norm(nome)] = true;
    cs.appendRow([ Utilities.getUuid(), nome, 'pendente', '', '', '', '', now, now ]);
    added++;
  });
  return json({ ok:true, added:added });
}

function handleUpdateGuest(data){
  var id = String(data.convidado_id || '');
  if (!id) return json({ ok:false, error:'sem id' });

  var cs = sheet('convidados', H_CONV);
  var rows = cs.getDataRange().getValues();
  var i = findRow(rows, 0, id);
  if (i < 1) return json({ ok:false, error:'convidado nao encontrado' });

  var cur = rows[i];
  var nome   = data.nome   != null ? clip(String(data.nome).trim())   : cur[1];
  var status = data.status != null ? String(data.status)              : cur[2];
  if (['pendente','vai','nao'].indexOf(status) < 0) status = cur[2];
  var adultos  = data.adultos  != null ? cleanNames(data.adultos).join(', ')  : cur[3];
  var criancas = data.criancas != null ? cleanNames(data.criancas).join(', ') : cur[4];
  var nota     = data.nota     != null ? clip(String(data.nota).trim())        : cur[5];

  cs.getRange(i+1, 2, 1, 8).setValues([[ nome, status, adultos, criancas, nota, cur[6], cur[7] || new Date(), new Date() ]]);
  return json({ ok:true });
}

function handleDeleteGuest(data){
  var id = String(data.convidado_id || '');
  var cs = sheet('convidados', H_CONV);
  var rows = cs.getDataRange().getValues();
  var i = findRow(rows, 0, id);
  if (i < 1) return json({ ok:false, error:'convidado nao encontrado' });
  cs.deleteRow(i+1);
  return json({ ok:true });
}

/** Liga uma resposta a um convidado existente, ou cria um novo convidado a partir dela. */
function handleLinkResponse(data){
  var respId = String(data.resposta_id || '');
  var rs = sheet('respostas', H_RESP);
  var rrows = rs.getDataRange().getValues();
  var ri = findRow(rrows, 0, respId);
  if (ri < 1) return json({ ok:false, error:'resposta nao encontrada' });
  var r = rrows[ri];
  var status = r[2], adultos = r[3], criancas = r[4], nota = r[8], nomeResp = r[7];
  var now = new Date();

  var cs = sheet('convidados', H_CONV);

  if (data.convidado_id === '__new' || !data.convidado_id){
    var nome = clip(String(data.nome || nomeResp || 'Sem nome').trim());
    var newId = Utilities.getUuid();
    cs.appendRow([ newId, nome, status, adultos, criancas, nota, respId, now, now ]);
    rs.getRange(ri+1, 10).setValue(newId);
    return json({ ok:true });
  }

  var crows = cs.getDataRange().getValues();
  var ci = findRow(crows, 0, String(data.convidado_id));
  if (ci < 1) return json({ ok:false, error:'convidado nao encontrado' });
  cs.getRange(ci+1, 3, 1, 7).setValues([[ status, adultos, criancas, nota, respId, crows[ci][7] || now, now ]]);
  rs.getRange(ri+1, 10).setValue(String(data.convidado_id));
  return json({ ok:true });
}

/* ============================ auxiliares ============================ */

function checkKey(key){
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!real && String(key || '') === String(real);
}

function withLock(fn){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return json({ ok:false, error:'ocupado, tenta de novo' }); }
  try { return fn(); } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** Devolve a folha; cria-a com cabeçalhos se não existir. */
function sheet(name, headers){
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); sh.appendRow(headers); }
  else if (sh.getLastRow() === 0){ sh.appendRow(headers); }
  return sh;
}

function getSpreadsheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  // Se o script for autónomo, guarda/recupera o id do Sheet nas propriedades.
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var created = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty('SHEET_ID', created.getId());
  return created;
}

function readSheet(name, headers){
  var sh = sheet(name, headers);
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++){
    var o = {}, row = values[i];
    if (!String(row[0] || '')) continue;
    for (var j = 0; j < headers.length; j++){
      var v = row[j];
      o[headers[j]] = (v instanceof Date) ? v.toISOString() : v;
    }
    out.push(o);
  }
  return out;
}

function findRow(rows, col, value){
  for (var i = 1; i < rows.length; i++) if (String(rows[i][col]) === String(value)) return i;
  return -1;
}

function cleanNames(arr){
  if (!Array.isArray(arr)) arr = String(arr || '').split(',');
  return arr.map(function(x){ return clip(String(x || '').trim()); }).filter(String).slice(0, 40);
}

function clip(s){ s = String(s == null ? '' : s); return s.length > MAX_LEN ? s.slice(0, MAX_LEN) : s; }

function norm(s){
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ aviso por email (opcional) ============================
 * Descomenta para receberes um email a cada nova resposta (complemento ao WhatsApp).
 * Em Propriedades do script, cria NOTIFY_EMAIL com o teu email.
 */
// function notifyEmail(status, nome, adultos, criancas, nota){
//   var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
//   if (!to) return;
//   var assunto = (status === 'vai' ? '✅ VAI' : '❌ Não vai') + ' – ' + (nome || 'Sem nome') + ' (Vitoria 4 anos)';
//   var corpo = 'Nome: ' + (nome || '-') +
//     '\nAdultos: ' + adultos.join(', ') +
//     '\nCrianças: ' + criancas.join(', ') +
//     '\nNota: ' + (nota || '-');
//   MailApp.sendEmail(to, assunto, corpo);
// }
