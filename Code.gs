/** Borcelle Clinic – backend. Script Properties: SPREADSHEET_ID (optional if bound), GEMINI_API_KEY (optional).
 *  Run setupDatabase() once, then setAdminPassword('your-strong-password') once. */
var TABLES = {
  Appointments: ['BookingID','CreatedAt','PatientName','Phone','WhatsApp','Email','Language','Gender','DateOfBirth','ServiceID','ServiceName','DoctorID','DoctorName','AppointmentDate','AppointmentTime','Reason','Notes','Status','PaymentStatus','PaymentMethod','CancellationReason','UpdatedAt'],
  Doctors: ['DoctorID','DoctorName','Specialty','Experience','Biography','ImageURL','AvailableDays','StartTime','EndTime','SlotDuration','Active'],
  Services: ['ServiceID','ServiceName','Description','ImageURL','Duration','Active'],
  Products: ['ProductID','ProductName','Category','Brand','Description','Price','ImageURL','Color','Material','Size','Availability','Active'],
  ContactMessages: ['MessageID','CreatedAt','Name','Phone','Email','Subject','Message','Status'],
  ClinicSettings: ['SettingKey','SettingValue'],
  BlockedDates: ['Date','Reason','CreatedAt'],
  AuditLogs: ['LogID','Timestamp','AdminAction','ReferenceID','Details'],
  FAQs: ['FAQID','QuestionEnglish','QuestionArabic','AnswerEnglish','AnswerArabic','Active'],
  Availability: ['AvailabilityID','DoctorID','DayOfWeek','StartTime','EndTime','BreakStart','BreakEnd','SlotDuration','Active']
};
var ADMIN_TABLES = {Doctors:'DoctorID',Services:'ServiceID',Products:'ProductID',FAQs:'FAQID',ClinicSettings:'SettingKey'};
var TZ = Session.getScriptTimeZone();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Borcelle Clinic – Eye Clinic & Vision Care')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function ok(data, message) { return {success: true, message: message || 'OK', data: data || {}}; }
function fail(message) { return {success: false, message: message, data: {}}; }
function safe(fn) { try { return fn(); } catch (e) { console.error(e); return fail(e.userMessage || 'Something went wrong. Please try again.'); } }
function userErr(m) { var e = new Error(m); e.userMessage = m; return e; }

// ---------- Database ----------
function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function createSheetIfMissing(name) {
  var ss = getSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]).setFontWeight('bold');
    sh.getRange(2, 1, 999, TABLES[name].length).setNumberFormat('@'); // keep dates/times as text
    sh.setFrozenRows(1);
  }
  return sh;
}
function setupDatabase() {
  Object.keys(TABLES).forEach(createSheetIfMissing);
  var s = getSpreadsheet();
  if (s.getSheetByName('Sheet1') && s.getSheets().length > 1) s.deleteSheet(s.getSheetByName('Sheet1'));
  var set = s.getSheetByName('ClinicSettings');
  if (set.getLastRow() < 2) {
    var d = [['ClinicName','Borcelle Clinic'],['Phone','+000 000 0000'],['WhatsApp','000000000000'],['Telegram','your_username'],['Email','clinic@example.com'],
      ['Address','Clinic address (placeholder)'],['MapsEmbedURL',''],['Website','www.example.com'],['OpeningHours','Sun–Thu 09:00–17:00'],
      ['ClosedDays','Fri,Sat'],['DefaultLanguage','ar'],['AutoConfirm','no'],['AdminEmail','']];
    set.getRange(2, 1, d.length, 2).setValues(d);
  }
  var sv = s.getSheetByName('Services');
  if (sv.getLastRow() < 2) sv.getRange(2, 1, 4, 6).setValues([
    ['S1','Comprehensive Eye Examination','Full vision and eye-health check (sample content).','','30','TRUE'],
    ['S2','Pediatric Eye Care','Eye care for children (sample content).','','30','TRUE'],
    ['S3','Contact Lens Fitting','Fitting and follow-up (sample content).','','30','TRUE'],
    ['S4','Glaucoma Screening','Eye pressure and optic nerve screening (sample content).','','30','TRUE']]);
  var dr = s.getSheetByName('Doctors');
  if (dr.getLastRow() < 2) dr.getRange(2, 1, 1, 11).setValues([['D1','Sample Doctor (replace)','Ophthalmologist','—','Sample profile. Replace with real clinic information.','','Sun,Mon,Tue,Wed,Thu','09:00','16:00','30','TRUE']]);
  var f = s.getSheetByName('FAQs');
  if (f.getLastRow() < 2) f.getRange(2, 1, 2, 6).setValues([
    ['F1','How can I book an appointment?','كيف يمكنني حجز موعد؟','Use the Book Appointment section: choose a service, doctor, date and time.','استخدم قسم حجز موعد: اختر الخدمة والطبيب والتاريخ والوقت.','TRUE'],
    ['F2','How do I pay?','كيف أدفع؟','Payment will be made at the clinic after your treatment or consultation.','يتم الدفع في العيادة بعد العلاج أو الاستشارة.','TRUE']]);
}
function rows_(name) {
  var v = getSpreadsheet().getSheetByName(name).getDataRange().getValues(), h = v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = String(r[i]); }); return o; });
}
function clean_(s, max) {
  s = String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // block formula injection
}
function append_(name, obj) {
  getSpreadsheet().getSheetByName(name).appendRow(TABLES[name].map(function (k) { return obj[k] == null ? '' : obj[k]; }));
}
function getClinicSettings() { var o = {}; rows_('ClinicSettings').forEach(function (r) { o[r.SettingKey] = r.SettingValue; }); return o; }
function isActive_(r) { return String(r.Active).toUpperCase() !== 'FALSE'; }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

// ---------- Public data ----------
function getServices() { return rows_('Services').filter(isActive_); }
function getDoctors() { return rows_('Doctors').filter(isActive_); }
function getProducts() { return rows_('Products').filter(isActive_); }
function getFAQs() { return rows_('FAQs').filter(isActive_); }
function getPublicData() {
  return safe(function () {
    var s = getClinicSettings(); delete s.AdminEmail;
    return ok({settings: s, services: getServices(), doctors: getDoctors(), products: getProducts(), faqs: getFAQs()});
  });
}

// ---------- Availability ----------
function toMin_(t) { var p = String(t).split(':'); return +p[0] * 60 + +p[1]; }
function fromMin_(m) { return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + m % 60).slice(-2); }
function slotsFor_(doctorId, date, appts) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw userErr('Invalid date.');
  var doc = rows_('Doctors').filter(function (d) { return d.DoctorID === doctorId && isActive_(d); })[0];
  if (!doc) throw userErr('Doctor not found.');
  var set = getClinicSettings(), dt = new Date(date + 'T00:00:00');
  var day = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][dt.getDay()];
  var closed = String(set.ClosedDays || '').split(',').map(function (x) { return x.trim(); });
  if (closed.indexOf(day) > -1 || doc.AvailableDays.split(',').map(function (x) { return x.trim(); }).indexOf(day) < 0) return [];
  if (rows_('BlockedDates').some(function (b) { return b.Date === date; })) return [];
  var step = +doc.SlotDuration || 30, taken = (appts || rows_('Appointments')).filter(function (a) {
    return a.DoctorID === doctorId && a.AppointmentDate === date && ['Cancelled','Rejected'].indexOf(a.Status) < 0; }).map(function (a) { return a.AppointmentTime; });
  var nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), out = [];
  for (var m = toMin_(doc.StartTime); m + step <= toMin_(doc.EndTime); m += step) {
    var t = fromMin_(m);
    if (taken.indexOf(t) < 0 && (date + ' ' + t) > nowStr) out.push(t);
  }
  return out;
}
function getAvailableSlots(doctorId, date) { return safe(function () { return ok({slots: slotsFor_(clean_(doctorId, 20), clean_(date, 10))}); }); }

// ---------- Booking ----------
function createAppointment(f) {
  return safe(function () {
    if (f.website) return ok({}, 'OK'); // honeypot
    rateLimit_('book', 10, 600);
    var name = clean_(f.name, 80), phone = clean_(f.phone, 20), wa = clean_(f.whatsapp || f.phone, 20), email = clean_(f.email, 100);
    if (name.length < 2) throw userErr('Please enter your full name.');
    if (!/^\+?[0-9 \-]{7,18}$/.test(phone)) throw userErr('Please enter a valid phone number.');
    if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw userErr('Please enter a valid email.');
    if (!f.consent) throw userErr('Please accept the consent.');
    var svc = getServices().filter(function (s) { return s.ServiceID === f.serviceId; })[0];
    var doc = getDoctors().filter(function (d) { return d.DoctorID === f.doctorId; })[0];
    if (!svc || !doc) throw userErr('Please select a service and a doctor.');
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var appts = rows_('Appointments');
      if (slotsFor_(doc.DoctorID, f.date, appts).indexOf(f.time) < 0) throw userErr('This time is no longer available. Please choose another.');
      var year = new Date().getFullYear(), n = appts.filter(function (a) { return a.BookingID.indexOf('EYE-' + year) === 0; }).length + 1;
      var id = 'EYE-' + year + '-' + ('0000' + n).slice(-4);
      var set = getClinicSettings(), status = set.AutoConfirm === 'yes' ? 'Confirmed' : 'Pending';
      append_('Appointments', {BookingID: id, CreatedAt: now_(), PatientName: name, Phone: phone, WhatsApp: wa, Email: email, Language: clean_(f.lang, 5),
        Gender: clean_(f.gender, 10), DateOfBirth: clean_(f.dob, 20), ServiceID: svc.ServiceID, ServiceName: svc.ServiceName, DoctorID: doc.DoctorID, DoctorName: doc.DoctorName,
        AppointmentDate: f.date, AppointmentTime: f.time, Reason: clean_(f.reason, 300), Notes: clean_(f.notes, 500), Status: status, PaymentStatus: 'Unpaid', UpdatedAt: now_()});
      var rec = {BookingID: id, PatientName: name, ServiceName: svc.ServiceName, DoctorName: doc.DoctorName, AppointmentDate: f.date, AppointmentTime: f.time, Status: status};
      try { sendBookingConfirmationEmail(rec, email, set); sendAdminBookingNotification(rec, set); } catch (e) { console.error('mail', e); }
      return ok(rec, 'Appointment created successfully');
    } finally { lock.releaseLock(); }
  });
}
function mail_(to, subject, body) { if (to) MailApp.sendEmail({to: to, subject: subject, body: body}); }
function sendBookingConfirmationEmail(r, email, set) {
  mail_(email, 'Booking ' + r.BookingID + ' – ' + (set.ClinicName || 'Clinic'), 'Dear ' + r.PatientName + ',\nYour appointment request (' + r.BookingID + ') for ' + r.ServiceName + ' with ' + r.DoctorName + ' on ' + r.AppointmentDate + ' at ' + r.AppointmentTime + ' is ' + r.Status + '.\nPayment will be made at the clinic after your treatment or consultation.');
}
function sendAdminBookingNotification(r, set) {
  mail_(set.AdminEmail || set.Email, 'New booking ' + r.BookingID, r.PatientName + ' – ' + r.ServiceName + ' – ' + r.AppointmentDate + ' ' + r.AppointmentTime);
}
function sendStatusUpdateNotification(a, status, set) {
  mail_(a.Email, 'Appointment ' + a.BookingID + ': ' + status, 'Your appointment on ' + a.AppointmentDate + ' at ' + a.AppointmentTime + ' is now: ' + status + '.');
}
function verifyAppointmentOwner(ref, phone) {
  var digits = function (s) { return String(s).replace(/\D/g, ''); };
  var a = rows_('Appointments').filter(function (x) { return x.BookingID === String(ref).trim().toUpperCase(); })[0];
  return a && digits(a.Phone) === digits(phone) && digits(phone).length > 5 ? a : null;
}
function getAppointmentByReference(ref, phone) {
  return safe(function () {
    rateLimit_('lookup', 10, 600);
    var a = verifyAppointmentOwner(ref, phone);
    if (!a) throw userErr('No booking found for this reference and phone number.');
    return ok({BookingID: a.BookingID, PatientName: a.PatientName, ServiceName: a.ServiceName, DoctorName: a.DoctorName, AppointmentDate: a.AppointmentDate, AppointmentTime: a.AppointmentTime, Status: a.Status});
  });
}
function setStatus_(id, status, reason, adminAction) {
  var sh = getSpreadsheet().getSheetByName('Appointments'), v = sh.getDataRange().getValues(), h = v[0], c = function (k) { return h.indexOf(k); };
  for (var i = 1; i < v.length; i++) if (v[i][0] === id) {
    sh.getRange(i + 1, c('Status') + 1).setValue(status);
    if (reason) sh.getRange(i + 1, c('CancellationReason') + 1).setValue(clean_(reason, 200));
    sh.getRange(i + 1, c('UpdatedAt') + 1).setValue(now_());
    var a = {}; h.forEach(function (k, j) { a[k] = String(v[i][j]); });
    audit_(adminAction, id, status);
    try { sendStatusUpdateNotification(a, status, getClinicSettings()); } catch (e) {}
    return true;
  }
  return false;
}
function cancelAppointment(ref, phone, reason) {
  return safe(function () {
    var a = verifyAppointmentOwner(ref, phone);
    if (!a) throw userErr('No booking found for this reference and phone number.');
    if (['Completed','Cancelled','Rejected','No Show'].indexOf(a.Status) > -1) throw userErr('This booking can no longer be cancelled.');
    setStatus_(a.BookingID, 'Cancelled', reason || 'Cancelled by patient', 'PatientCancel');
    return ok({}, 'Booking cancelled');
  });
}
function requestReschedule(ref, phone, note) {
  return safe(function () {
    var a = verifyAppointmentOwner(ref, phone);
    if (!a) throw userErr('No booking found for this reference and phone number.');
    audit_('RescheduleRequest', a.BookingID, clean_(note, 200));
    mail_(getClinicSettings().AdminEmail || getClinicSettings().Email, 'Reschedule request ' + a.BookingID, clean_(note, 300));
    return ok({}, 'Reschedule request sent to the clinic');
  });
}
function audit_(action, ref, details) { append_('AuditLogs', {LogID: Utilities.getUuid(), Timestamp: now_(), AdminAction: action, ReferenceID: ref, Details: details}); }
function rateLimit_(key, max, secs) {
  var c = CacheService.getScriptCache(), n = +(c.get('rl_' + key) || 0);
  if (n >= max) throw userErr('Too many attempts. Please wait a few minutes.');
  c.put('rl_' + key, String(n + 1), secs);
}

// ---------- Contact ----------
function submitContactMessage(f) {
  return safe(function () {
    if (f.website) return ok({});
    rateLimit_('contact', 10, 600);
    var name = clean_(f.name, 80), msg = clean_(f.message, 1000), email = clean_(f.email, 100);
    if (name.length < 2 || msg.length < 5) throw userErr('Please fill in your name and message.');
    if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw userErr('Please enter a valid email.');
    append_('ContactMessages', {MessageID: 'M' + Date.now(), CreatedAt: now_(), Name: name, Phone: clean_(f.phone, 20), Email: email, Subject: clean_(f.subject, 120), Message: msg, Status: 'New'});
    return ok({}, 'Message received');
  });
}

// ---------- Chatbot ----------
function askChatbot(question, lang) {
  return safe(function () {
    rateLimit_('chat', 40, 600);
    var q = clean_(question, 400), ar = lang === 'ar', set = getClinicSettings(), faqs = getFAQs();
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    var urgent = /sudden|loss of vision|injur|severe pain|chemical|فقدان|إصابة|ألم شديد|مفاجئ/i.test(q);
    if (urgent) return ok({answer: ar ? 'قد تكون هذه حالة طارئة. يرجى التوجه فوراً إلى أقرب مركز طوارئ أو طبيب عيون.' : 'This may be urgent. Please seek immediate medical attention at an emergency department or from an eye specialist.'});
    if (key) {
      var ctx = 'Clinic: ' + set.ClinicName + '. Hours: ' + set.OpeningHours + '. Address: ' + set.Address + '. Phone: ' + set.Phone + '. Services: ' + getServices().map(function (s) { return s.ServiceName; }).join(', ') +
        '. FAQs: ' + faqs.map(function (f) { return f.QuestionEnglish + ' ' + f.AnswerEnglish; }).join(' | ') + '. Payment is at the clinic after treatment.';
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent', {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: {'x-goog-api-key': key}, payload: JSON.stringify({systemInstruction: {parts: [{text: 'You are a clinic information assistant. Use only this data: ' + ctx + ' Reply in the user\'s language, briefly. Never diagnose, prescribe, or give treatment advice; for medical concerns recommend seeing an eye specialist.'}]}, contents: [{parts: [{text: q}]}]})});
      if (res.getResponseCode() === 200) {
        var t = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        return ok({answer: t});
      }
    }
    var words = q.toLowerCase().split(/\s+/), best = null, score = 0;
    faqs.forEach(function (f) { var s = 0, t = (f.QuestionEnglish + ' ' + f.QuestionArabic).toLowerCase(); words.forEach(function (w) { if (w.length > 2 && t.indexOf(w) > -1) s++; }); if (s > score) { score = s; best = f; } });
    if (best) return ok({answer: ar ? best.AnswerArabic : best.AnswerEnglish});
    return ok({answer: ar ? 'يمكنكم التواصل مع العيادة على ' + set.Phone : 'Please contact the clinic on ' + set.Phone + ' or use the booking form.'});
  });
}

// ---------- Admin auth ----------
function hash_(salt, pw) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function setAdminPassword(pw) { // run once from the editor
  if (!pw || pw.length < 10) throw new Error('Use at least 10 characters.');
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ADMIN_SALT: salt, ADMIN_HASH: hash_(salt, pw)});
}
function adminLogin(password) {
  return safe(function () {
    var p = PropertiesService.getScriptProperties(), c = CacheService.getScriptCache(), fails = +(c.get('loginFails') || 0);
    if (fails >= 5) throw userErr('Too many failed attempts. Try again in 15 minutes.');
    if (!p.getProperty('ADMIN_HASH') || hash_(p.getProperty('ADMIN_SALT'), String(password)) !== p.getProperty('ADMIN_HASH')) {
      c.put('loginFails', String(fails + 1), 900); throw userErr('Invalid credentials.');
    }
    var token = Utilities.getUuid() + Utilities.getUuid(); c.put('sess_' + token, '1', 3600);
    return ok({token: token});
  });
}
function validateAdminSession(token) { return !!token && CacheService.getScriptCache().get('sess_' + token) === '1'; }
function adminLogout(token) { CacheService.getScriptCache().remove('sess_' + token); return ok({}); }
function guard_(token) { if (!validateAdminSession(token)) throw userErr('Session expired. Please log in again.'); }

// ---------- Admin operations ----------
function getDashboardStats(token) {
  return safe(function () {
    guard_(token);
    var a = rows_('Appointments'), by = function (s) { return a.filter(function (x) { return x.Status === s; }).length; }, svc = {};
    a.forEach(function (x) { svc[x.ServiceName] = (svc[x.ServiceName] || 0) + 1; });
    var phones = {}; a.forEach(function (x) { phones[x.Phone] = 1; });
    return ok({total: a.length, pending: by('Pending'), confirmed: by('Confirmed'), completed: by('Completed'), cancelled: by('Cancelled'), patients: Object.keys(phones).length,
      products: getProducts().length, newMessages: rows_('ContactMessages').filter(function (m) { return m.Status === 'New'; }).length, services: svc});
  });
}
function getAllAppointments(token) { return safe(function () { guard_(token); return ok({items: rows_('Appointments').reverse()}); }); }
function updateAppointmentStatus(token, id, status, reason) {
  return safe(function () {
    guard_(token);
    if (['Pending','Confirmed','Completed','Cancelled','Rejected','No Show'].indexOf(status) < 0) throw userErr('Invalid status.');
    if (!setStatus_(clean_(id, 30), status, reason, 'Status:' + status)) throw userErr('Booking not found.');
    return ok({}, 'Updated');
  });
}
function getContactMessages(token) { return safe(function () { guard_(token); return ok({items: rows_('ContactMessages').reverse()}); }); }
function getAuditLogs(token) { return safe(function () { guard_(token); return ok({items: rows_('AuditLogs').reverse().slice(0, 200)}); }); }
// Generic CRUD for Doctors, Services, Products, FAQs, ClinicSettings (covers add/update/delete/deactivate)
function adminSaveRecord(token, table, rec) {
  return safe(function () {
    guard_(token);
    var idCol = ADMIN_TABLES[table]; if (!idCol) throw userErr('Not allowed.');
    var sh = getSpreadsheet().getSheetByName(table), v = sh.getDataRange().getValues(), h = v[0];
    rec = rec || {};
    if (!rec[idCol]) rec[idCol] = table.charAt(0) + Date.now();
    var line = h.map(function (k) { return clean_(rec[k], 1000); });
    for (var i = 1; i < v.length; i++) if (v[i][0] === rec[idCol]) { sh.getRange(i + 1, 1, 1, h.length).setValues([line]); audit_('Update ' + table, rec[idCol], ''); return ok({id: rec[idCol]}); }
    sh.appendRow(line); audit_('Add ' + table, rec[idCol], ''); return ok({id: rec[idCol]});
  });
}
function adminDeleteRecord(token, table, id) {
  return safe(function () {
    guard_(token);
    if (!ADMIN_TABLES[table]) throw userErr('Not allowed.');
    var sh = getSpreadsheet().getSheetByName(table), v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) if (v[i][0] === id) { sh.deleteRow(i + 1); audit_('Delete ' + table, id, ''); return ok({}); }
    throw userErr('Record not found.');
  });
}
function updateContactMessageStatus(token, id, status) {
  return safe(function () {
    guard_(token);
    var sh = getSpreadsheet().getSheetByName('ContactMessages'), v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) if (v[i][0] === id) { sh.getRange(i + 1, 8).setValue(status === 'Read' ? 'Read' : 'New'); return ok({}); }
    throw userErr('Not found.');
  });
}
