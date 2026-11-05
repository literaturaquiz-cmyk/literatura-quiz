// Google Apps Script para Literatura Quiz. Pega este archivo en script.google.com.
const TEACHER_USER = 'Marcia';
const TEACHER_PASSWORD = '1221*';
const RESULTS_SHEET = 'Resultados';
const QUIZZES_SHEET = 'Trivias';
function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = getSheet(ss, RESULTS_SHEET, ['Nombre','Grado','Movimiento','Puntaje','Porcentaje','Fecha']);
    const quizzes = getSheet(ss, QUIZZES_SHEET, ['Título','Datos JSON']);
    let out;
    if (data.action === 'login') {
      requireTeacher(data); out = {ok:true};
    } else if (data.action === 'getQuizzes') {
      const qvals = quizzes.getDataRange().getValues();
      const saved = qvals.slice(1).filter(r=>r[1]).map(r=>{try{return JSON.parse(r[1]);}catch(_){return null;}}).filter(Boolean);
      out = {ok:true,quizzes:saved};
    } else if (data.action === 'submit') {
      const r = data.result || {};
      if (!r.name || !r.grade || !r.level || !Number.isFinite(Number(r.score))) throw new Error('Faltan datos del resultado.');
      results.appendRow([r.name,r.grade,r.level,Number(r.score),Number(r.percentage),r.date || new Date().toLocaleString('es-CO')]);
      out = {ok:true};
    } else if (data.action === 'list') {
      requireTeacher(data);
      const vals = results.getDataRange().getValues();
      const rows = vals.slice(1).filter(r=>r[0]).map(r=>({name:r[0],grade:r[1],level:r[2],score:r[3],percentage:r[4],date:r[5]}));
      const qvals = quizzes.getDataRange().getValues();
      const saved = qvals.slice(1).filter(r=>r[1]).map(r=>{try{return JSON.parse(r[1]);}catch(_){return null;}}).filter(Boolean);
      out = {ok:true,results:rows,quizzes:saved};
    } else if (data.action === 'addQuiz') {
      requireTeacher(data);
      const q = data.quiz;
      if (!q || !q.title || !Array.isArray(q.questions) || q.questions.length !== 15) throw new Error('Cada trivia debe tener exactamente 15 preguntas.');
      quizzes.appendRow([q.title,JSON.stringify(q)]); out = {ok:true};
    } else { throw new Error('Acción no reconocida.'); }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
function getSheet(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh){sh=ss.insertSheet(name);sh.appendRow(headers);}return sh;}

function requireTeacher(data){if(data.username!==TEACHER_USER || data.password!==TEACHER_PASSWORD) throw new Error('Acceso docente no autorizado.');}
