// Grupo acadêmico — backend para pesquisa Mariana Ferrer
// 1) Crie uma planilha Google e abra Extensões > Apps Script.
// 2) Cole este código. Em Configurações do projeto > Propriedades do script,
//    crie a propriedade ACCESS_PASSWORD com o valor PÃE2026.
// 3) Implante como Aplicativo da Web: executar como você; acesso: qualquer pessoa.
// 4) Copie a URL /exec para a constante API_URL do HTML.
// A planilha vinculada será usada para armazenar respostas anônimas.

const SHEET_NAME = 'Respostas';
function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (payload.action !== 'submit' || !payload.answers) return json_({ok:false, message:'Solicitação inválida.'});
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Data/hora','Tema relevante','Conhecia caso/lei','Como evitar revitimização']);
    }
    const a = payload.answers;
    sheet.appendRow([new Date(), safeCell_(a.q1), safeCell_(a.q2), safeCell_(a.q3)]);
    return json_({ok:true});
  } catch (err) {
    return json_({ok:false, message:'Não foi possível registrar a resposta.'});
  }
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '').replace(/[^a-zA-Z0-9_$]/g, '');
  let result;
  const expected = PropertiesService.getScriptProperties().getProperty('ACCESS_PASSWORD');
  if (p.action !== 'results') result = {ok:false, message:'Ação inválida.'};
  else if (!expected || p.password !== expected) result = {ok:false, message:'Senha incorreta.'};
  else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const rows = sheet && sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,4).getDisplayValues() : [];
    const summary = {q1:{},q2:{},q3:{}};
    rows.forEach(r => { [['q1',r[1]],['q2',r[2]],['q3',r[3]]].forEach(([k,v])=>{ if(v) summary[k][v]=(summary[k][v]||0)+1; }); });
    result = {ok:true, total:rows.length, summary:summary};
  }
  const output = JSON.stringify(result);
  if (callback) return ContentService.createTextOutput(callback+'('+output+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
}
function safeCell_(value) { return String(value || '').replace(/[\r\n]/g,' ').slice(0,500); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
