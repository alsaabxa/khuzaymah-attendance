/**
 * منصة رصد الغياب
 * ابن خلدون الابتدائية وخزيمة بن ثابت المتوسطة
 * المدير: الأستاذ صافي اللحياني
 * النوع: طلاب
 *
 * طريقة الاستخدام:
 * 1) اربط هذا المشروع بملف Google Sheet الجديد.
 * 2) الصق الكود وشغّل setup_ مرة واحدة.
 * 3) انشره Web App: Execute as Me + Anyone.
 * 4) ضع رابط /exec في index.html داخل SCRIPT_URL.
 */

const SCHOOL_NAME = 'ابن خلدون الابتدائية وخزيمة بن ثابت المتوسطة';
const SCHOOL_TYPE = 'طلاب';
const SCHOOL_PRINCIPAL = 'الأستاذ صافي اللحياني';

const SHEETS = {
  classes: ['Classes', ['id','number','name']],
  students: ['Students', ['id','classId','number','name']],
  staff: ['Staff', ['id','name','role','username','password','classes']],
  attendance: ['Attendance', ['id','classId','date','teacher','time','studentId','studentName','status']],
  teacherNames: ['TeacherNames', ['name']]
};

function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(k => {
    const [name, headers] = SHEETS[k];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  });
  return 'تم تجهيز جداول المنصة';
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '');
  if (action === 'bootstrap' || !action) return json_(bootstrap_());
  return json_({ok:false,error:'Unknown action'});
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '');
    const payload = body.payload || {};
    const result = handleAction_(action, payload);
    return json_(result);
  } catch (err) {
    return json_({ok:false,error:String(err)});
  }
}

function bootstrap_() {
  setup_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classes = readObjects_(ss.getSheetByName('Classes')).map(r => ({
    id: str_(r.id), number: str_(r.number), name: str_(r.name)
  })).filter(r => r.id);

  const students = readObjects_(ss.getSheetByName('Students')).map(r => ({
    id: str_(r.id), classId: str_(r.classId), number: str_(r.number), name: str_(r.name)
  })).filter(r => r.id);

  const staff = readObjects_(ss.getSheetByName('Staff')).map(r => ({
    id: str_(r.id), name: str_(r.name), role: str_(r.role),
    username: str_(r.username), password: str_(r.password),
    classes: str_(r.classes) ? str_(r.classes).split(',').filter(Boolean) : []
  })).filter(r => r.id);

  const attendance = readObjects_(ss.getSheetByName('Attendance')).map(r => ({
    id: str_(r.id), classId: str_(r.classId), date: str_(r.date),
    teacher: str_(r.teacher), time: str_(r.time),
    studentId: str_(r.studentId), studentName: str_(r.studentName), status: str_(r.status)
  })).filter(r => r.classId && r.date && r.studentId);

  const teacherNames = readObjects_(ss.getSheetByName('TeacherNames'))
    .map(r => str_(r.name)).filter(Boolean);

  return {ok:true, schoolName:SCHOOL_NAME, schoolType:SCHOOL_TYPE, principal:SCHOOL_PRINCIPAL,
    classes, students, staff, attendance, teacherNames};
}

function handleAction_(action, p) {
  setup_();
  switch(action) {
    case 'saveAttendance': return saveAttendance_(p);
    case 'deleteClass': return deleteClass_(p);
    case 'deleteToday': return deleteToday_(p);
    case 'deleteStudent': return deleteById_('Students', p.id);
    case 'saveStudent': return upsertStudent_(p);
    case 'saveClass': return upsertClass_(p);
    case 'saveClassWithStudents': return saveClassWithStudents_(p);
    case 'saveTeacherName': return saveTeacherName_(p);
    case 'deleteTeacherName': return deleteTeacherName_(p);
    case 'saveStaff': return upsertStaff_(p);
    case 'deleteStaff': return deleteById_('Staff', p.id);
    default: return {ok:false,error:'Unknown action: '+action};
  }
}

function saveAttendance_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
  const rows = Array.isArray(p.records) ? p.records : [];
  if (!p.id) p.id = String(p.classId || '')+'|'+String(p.date || '');
  deleteRows_(sh, r => str_(r[0]) === str_(p.id));
  if (rows.length) {
    const values = rows.map(r => [p.id,p.classId,p.date,p.teacher || '',p.time || '',r.studentId || '',r.studentName || '',r.status || '']);
    sh.getRange(sh.getLastRow()+1,1,values.length,8).setValues(values);
  }
  return {ok:true};
}

function deleteClass_(p) {
  const id = str_(p.id);
  deleteRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Classes'), r => str_(r[0]) === id);
  deleteRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students'), r => str_(r[1]) === id);
  deleteRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance'), r => str_(r[1]) === id);
  return {ok:true};
}

function deleteToday_(p) {
  const classId = str_(p.classId), date = str_(p.date);
  deleteRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance'),
    r => str_(r[1]) === classId && str_(r[2]) === date);
  return {ok:true};
}

function upsertClass_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Classes');
  upsertRow_(sh, p.id, [p.id,p.number || '',p.name || ''], 1);
  return {ok:true};
}

function saveClassWithStudents_(p) {
  upsertClass_(p.classObj || {});
  const list = Array.isArray(p.students) ? p.students : [];
  list.forEach(s => upsertStudent_(s));
  return {ok:true};
}

function upsertStudent_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  upsertRow_(sh, p.id, [p.id,p.classId || '',p.number || '',p.name || ''], 1);
  return {ok:true};
}

function upsertStaff_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff');
  const classes = Array.isArray(p.classes) ? p.classes.join(',') : String(p.classes || '');
  upsertRow_(sh, p.id, [p.id,p.name || '',p.role || '',p.username || '',p.password || '',classes], 1);
  return {ok:true};
}

function saveTeacherName_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TeacherNames');
  const name = str_(p.name);
  if (!name) return {ok:false,error:'اسم المعلم فارغ'};
  const names = readObjects_(sh).map(r => str_(r.name));
  if (names.indexOf(name) < 0) sh.appendRow([name]);
  return {ok:true};
}

function deleteTeacherName_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TeacherNames');
  deleteRows_(sh, r => str_(r[0]) === str_(p.name));
  return {ok:true};
}

function deleteById_(sheetName, id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  deleteRows_(sh, r => str_(r[0]) === str_(id));
  return {ok:true};
}

function upsertRow_(sh, id, values, idColumn) {
  const target = str_(id);
  if (!target) throw new Error('Missing id');
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2,idColumn,last-1,1).getDisplayValues().flat();
    const idx = ids.findIndex(x => str_(x) === target);
    if (idx >= 0) {
      sh.getRange(idx+2,1,1,values.length).setValues([values]);
      return;
    }
  }
  sh.appendRow(values);
}

function deleteRows_(sh, predicate) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const values = sh.getRange(2,1,last-1,sh.getLastColumn()).getDisplayValues();
  for (let i=values.length-1;i>=0;i--) {
    if (predicate(values[i])) sh.deleteRow(i+2);
  }
}

function readObjects_(sh) {
  const last = sh.getLastRow(), cols = sh.getLastColumn();
  if (last < 2 || cols < 1) return [];
  const headers = sh.getRange(1,1,1,cols).getDisplayValues()[0];
  const values = sh.getRange(2,1,last-1,cols).getDisplayValues();
  return values.map(row => {
    const o = {};
    headers.forEach((h,i)=>o[h]=row[i]);
    return o;
  });
}

function str_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
