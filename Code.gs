/**
 * Configuration
 */
const SPREADSHEET_ID = '1VbPLmAyfxi5zy2JrQU3QgaqWiGRfxCGWcw5g92z19XM';

/**
 * INITIALIZATION
 */
function initProject() {
  const ss = getSS();
  let settingSheet = ss.getSheetByName('Setting');
  if (!settingSheet) settingSheet = ss.insertSheet('Setting');
  settingSheet.clear();
  settingSheet.appendRow(['Parameter', 'Value', 'Keterangan']);
  settingSheet.appendRow(['API_KEY', '9ed61aa8151df31bd0f9718b82f067462938f0a5cd3cfd312dac18069d34019b', 'API Key']);
  settingSheet.appendRow(['SENDER', '6285187232455', 'Nomor WA']);
  settingSheet.appendRow(['TEMPLATE_MASUK', '*Assalaamualaikum*\n{{nama}} {{kelas}} BERHASIL Masuk jam {{waktu}}.', 'Template']);
  settingSheet.appendRow(['TEMPLATE_PULANG', '*Assalaamualaikum*\n{{nama}} {{kelas}} BERHASIL Pulang jam {{waktu}}.', 'Template']);
  settingSheet.appendRow(['NAMA_LEMBAGA', 'MI Nurul Islam Gedangmas', 'Lembaga']);
  settingSheet.appendRow(['URL_LOGO', 'https://cdn-icons-png.flaticon.com/512/2859/2859135.png', 'Logo']);
  
  let logSheet = ss.getSheetByName('Log Presensi');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log Presensi');
    logSheet.appendRow(['Timestamp', 'Nama', 'Kelas', 'Status', 'Keterangan', 'Jenis']);
  }
}

/**
 * UTILS
 */
function getSS() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function getSettings() {
  const sheet = getSS().getSheetByName('Setting');
  if (!sheet) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
  const s = {};
  data.forEach(r => { if(r[0]) s[r[0]] = r[1]; });
  return s;
}
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }
function getDirectDriveLink(u) {
  if (!u || u.indexOf('drive.google.com') === -1) return u || '';
  const id = u.match(/[-\w]{25,}/);
  return id ? `https://drive.google.com/thumbnail?id=${id[0]}&sz=w500` : u;
}

/**
 * WEB APP
 */
function doGet() {
  const s = getSettings();
  const t = HtmlService.createTemplateFromFile('index');
  t.madrasahName = s.NAMA_LEMBAGA || 'Presensi Digital';
  t.madrasahLogo = getDirectDriveLink(s.URL_LOGO);
  return t.evaluate().setTitle(t.madrasahName).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * DATA FETCHING
 */
function getStudentData() {
  const sheet = getSS().getSheetByName('Data Siswa');
  if (!sheet) return [];
  const rows = sheet.getLastRow();
  if (rows < 2) return [];
  const data = sheet.getRange(2, 1, rows - 1, 4).getValues();
  return data.map(r => ({ nisn: r[0], nama: r[1], kelas: String(r[2] || '?').trim(), waOrtu: r[3] })).filter(x => x.nama);
}

function checkAlreadyAttended(name, dateStr) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Log Presensi');
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  
  // Check the last 1000 rows for better performance
  const data = sheet.getRange(Math.max(2, lastRow - 1000), 1, Math.min(lastRow - 1, 1001), 2).getValues();
  return data.some(r => {
    if (!r[0]) return false;
    const d = Utilities.formatDate(new Date(r[0]), "GMT+7", "yyyy-MM-dd");
    return d === dateStr && r[1] === name;
  });
}

function submitAttendance(d) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Log Presensi');
    
    const subDateStr = d.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    
    // Check Duplicates
    const duplicates = [];
    d.students.forEach(st => {
      if (checkAlreadyAttended(st.nama, subDateStr)) {
        duplicates.push(st.nama);
      }
    });

    if (duplicates.length > 0) {
      return { 
        success: false, 
        alreadyAttended: true,
        message: duplicates.length === 1 ? `${duplicates[0]} sudah absen hari ini.` : `Siswa berikut sudah absen: ${duplicates.join(', ')}`
      };
    }
    
    // Handle Date
    let ts;
    let allowWa = true;
    if (d.date) {
      ts = new Date(d.date);
      // Set to current time if it's today, otherwise keep as is (midnight)
      const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
      const subDateStr = d.date; // yyyy-MM-dd
      if (subDateStr !== todayStr) {
        allowWa = false;
      } else {
        ts = new Date(); // Use current time for today
      }
    } else {
      ts = new Date();
    }
    
    const s = getSettings();
    const all = getStudentData();
    const map = {};
    all.forEach(x => map[x.nama] = x.waOrtu);
    
    const rows = d.students.map(x => [ts, x.nama, d.kelas, x.status, d.keterangan || '-', d.jenis]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    
    if (allowWa && (d.status === 'Hadir' || d.students.some(x => x.status === 'Hadir'))) {
      d.students.forEach(x => {
        if (x.status === 'Hadir' && map[x.nama]) {
          sendWa(s, map[x.nama], x.nama, x.status, d.kelas, d.jenis);
        }
      });
    }
    return { success: true, message: 'Berhasil!' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function sendWa(s, to, n, st, k, j) {
  if (!to || !s.API_KEY || !s.SENDER) return;
  const time = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  const tk = (j === 'Pulang') ? 'TEMPLATE_PULANG' : 'TEMPLATE_MASUK';
  let msg = s[tk] || "";
  msg = msg.replace(/{{nama}}/g, n).replace(/{{kelas}}/g, k).replace(/{{status}}/g, st).replace(/{{waktu}}/g, time);
  const url = `https://gateway.pdmhadirq.cloud/api/send-message?api_key=${s.API_KEY}&sender=${s.SENDER}&number=${to}&message=${encodeURIComponent(msg)}`;
  UrlFetchApp.fetch(url, { muteHttpExceptions: true });
}

function getAttendanceSummary() {
  const sheet = getSS().getSheetByName('Log Presensi');
  if (!sheet) return [];
  const rows = sheet.getLastRow();
  if (rows < 2) return [];
  const data = sheet.getRange(Math.max(2, rows - 49), 1, Math.min(rows - 1, 50), 6).getValues();
  return data.map(r => ({
    timestamp: Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM HH:mm"),
    nama: r[1], kelas: r[2], status: r[3], jenis: r[5]
  })).reverse();
}

/**
 * REPORTING LOGIC
 */
function getReportData(type, dateStr, month, year, className) {
  const logSheet = getSS().getSheetByName('Log Presensi');
  const studentData = getStudentData().filter(s => s.kelas === className);
  if (studentData.length === 0) return [];

  const logs = logSheet.getLastRow() > 1 ? logSheet.getRange(2, 1, logSheet.getLastRow()-1, 6).getValues() : [];
  
  if (type === 'daily') {
    // Filter logs for selected date and class
    const targetDate = dateStr; // yyyy-MM-dd
    const dailyLogs = {};
    logs.forEach(r => {
      const d = Utilities.formatDate(new Date(r[0]), "GMT+7", "yyyy-MM-dd");
      if (d === targetDate && r[2] === className) {
        dailyLogs[r[1] + r[5]] = r[3]; // Name + Type (Masuk/Pulang)
      }
    });

    return studentData.map(s => ({
      nama: s.nama,
      masuk: dailyLogs[s.nama + 'Masuk'] || '-',
      pulang: dailyLogs[s.nama + 'Pulang'] || '-'
    }));
  } else {
    // Monthly Recap
    const targetMonth = parseInt(month); // 1-12
    const targetYear = parseInt(year);
    const summary = {};
    
    studentData.forEach(s => {
      summary[s.nama] = { H: 0, I: 0, S: 0, A: 0 };
    });

    logs.forEach(r => {
      const date = new Date(r[0]);
      if (date.getMonth() + 1 === targetMonth && date.getFullYear() === targetYear && r[2] === className && r[5] === 'Masuk') {
        const name = r[1];
        const status = r[3];
        if (summary[name]) {
          if (status === 'Hadir') summary[name].H++;
          else if (status === 'Izin') summary[name].I++;
          else if (status === 'Sakit') summary[name].S++;
          else if (status === 'Alpha') summary[name].A++;
        }
      }
    });

    return Object.keys(summary).map(name => ({
      nama: name,
      ...summary[name]
    }));
  }
}
