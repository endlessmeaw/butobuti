// ================================================================
//  LOTTOAI PRO — Google Apps Script Backend
//  เวอร์ชัน: 1.0
//  สร้างเมื่อ: 2026-04-19
//  ใช้ Google Sheets เป็นฐานข้อมูล
// ================================================================

// ╔═══════════════════════════════════════════╗
// ║  🛠️ ตั้งค่า — รันฟังก์ชัน setupSheets()   ║
// ║  ก่อนใช้งานครั้งแรก เพื่อสร้าง Sheet       ║
// ╚═══════════════════════════════════════════╝

/**
 * สร้าง Sheet "Users" และ "History" อัตโนมัติ
 * ⚡ รันฟังก์ชันนี้ครั้งเดียวก่อน Deploy
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // สร้าง Sheet "Users"
  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['username', 'password', 'created_at']);
    usersSheet.getRange('1:1').setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    usersSheet.setColumnWidths(1, 3, 200);
    usersSheet.setFrozenRows(1);
  }

  // สร้าง Sheet "History"
  let historySheet = ss.getSheetByName('History');
  if (!historySheet) {
    historySheet = ss.insertSheet('History');
    historySheet.appendRow(['username', 'type', 'mode', 'input', 'result', 'time', 'combosTwo', 'combosThree']);
    historySheet.getRange('1:1').setFontWeight('bold').setBackground('#0ea5e9').setFontColor('#ffffff');
    historySheet.setColumnWidths(1, 8, 160);
    historySheet.setFrozenRows(1);
  }

  Logger.log('✅ Setup เสร็จสมบูรณ์! สร้าง Sheet "Users" และ "History" แล้ว');
}


// ================================================================
//  📡 MAIN HANDLERS — รับ Request จาก Frontend
// ================================================================

/**
 * รับ POST requests (register, login, analyze)
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case 'register':
        return register_(body.data);
      case 'login':
        return login_(body.data);
      case 'analyze':
        return analyze_(body.data, body.username);
      default:
        return jsonResponse_({ status: 'error', message: 'ไม่พบ Action ที่ร้องขอ' });
    }
  } catch (err) {
    return jsonResponse_({ status: 'error', message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

/**
 * รับ GET requests (getHistory)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'getHistory') {
      return getHistory_(e.parameter.username);
    }

    return jsonResponse_({ status: 'error', message: 'ไม่พบ Action ที่ร้องขอ' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

/**
 * Helper: สร้าง JSON Response
 */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ================================================================
//  🔐 ระบบสมาชิก — สมัคร / เข้าสู่ระบบ
// ================================================================

/**
 * สมัครสมาชิก
 */
function register_(data) {
  if (!data.username || !data.password) {
    return jsonResponse_({ status: 'error', message: 'กรุณากรอกข้อมูลให้ครบ' });
  }

  const username = String(data.username).trim();
  const password = String(data.password).trim();

  if (username.length < 3) {
    return jsonResponse_({ status: 'error', message: 'ชื่อผู้ใช้ต้องมีอย่างน้อย 3 ตัวอักษร' });
  }
  if (password.length < 4) {
    return jsonResponse_({ status: 'error', message: 'รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const users = sheet.getDataRange().getValues();

  // ตรวจว่า username ซ้ำไหม
  for (let i = 1; i < users.length; i++) {
    if (String(users[i][0]).toLowerCase() === username.toLowerCase()) {
      return jsonResponse_({ status: 'error', message: 'ชื่อผู้ใช้นี้ถูกใช้แล้ว' });
    }
  }

  // Hash รหัสผ่าน
  const hashed = hashPassword_(password);
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

  sheet.appendRow([username, hashed, now]);

  return jsonResponse_({ status: 'success', message: 'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ' });
}

/**
 * เข้าสู่ระบบ
 */
function login_(data) {
  if (!data.username || !data.password) {
    return jsonResponse_({ status: 'error', message: 'กรุณากรอกข้อมูลให้ครบ' });
  }

  const username = String(data.username).trim();
  const password = String(data.password).trim();
  const hashed = hashPassword_(password);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const users = sheet.getDataRange().getValues();

  for (let i = 1; i < users.length; i++) {
    if (String(users[i][0]) === username && String(users[i][1]) === hashed) {
      return jsonResponse_({ status: 'success', username: username });
    }
  }

  return jsonResponse_({ status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
}

/**
 * Hash password ด้วย SHA-256
 */
function hashPassword_(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return raw.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}


// ================================================================
//  🎯 ระบบวิเคราะห์เลขเด่น — 75 สูตรหลังบ้าน
// ================================================================

/**
 * วิเคราะห์เลขเด่น
 * @param {Object} formData - ข้อมูลจาก form (lotteryType, inputNumbers, engineMode, numStates)
 * @param {String} username - ชื่อผู้ใช้
 */
function analyze_(formData, username) {
  const inputRaw = String(formData.inputNumbers || '000000');
  const input = inputRaw.padStart(6, '0').slice(0, 6);
  const mode = formData.engineMode || 'A';
  const numStates = formData.numStates || {};
  const lotteryType = formData.lotteryType || 'รัฐบาลไทย';

  // แยก digit
  const digits = input.split('').map(Number);

  // สร้าง list บังคับ/แบน
  const forced = [];
  const banned = [];
  for (let i = 0; i <= 9; i++) {
    const state = numStates[String(i)];
    if (state === 'force') forced.push(i);
    if (state === 'ban') banned.push(i);
  }

  // ================================================
  //  รัน 75 สูตร → ได้คะแนนแต่ละเลข 0-9
  // ================================================
  const scores = runAllFormulas_(digits, mode);

  // ปรับคะแนนตาม Force / Ban
  for (let i = 0; i <= 9; i++) {
    if (banned.includes(i)) scores[i] = -99999;
    if (forced.includes(i)) scores[i] = 99999;
  }

  // เรียงลำดับตามคะแนน
  const ranked = [];
  for (let i = 0; i <= 9; i++) {
    ranked.push({ digit: i, score: scores[i] });
  }
  ranked.sort(function (a, b) { return b.score - a.score; });

  // เลือกเลขเด่น 6 ตัว
  var prediction = [];

  // เพิ่มเลขบังคับก่อน
  for (var fi = 0; fi < forced.length; fi++) {
    if (prediction.indexOf(forced[fi]) === -1 && prediction.length < 6) {
      prediction.push(forced[fi]);
    }
  }

  // เติมเลขจากคะแนนสูงสุด
  for (var ri = 0; ri < ranked.length; ri++) {
    if (prediction.length >= 6) break;
    var d = ranked[ri].digit;
    if (prediction.indexOf(d) === -1 && banned.indexOf(d) === -1) {
      prediction.push(d);
    }
  }

  // กันกรณีเลขไม่พอ (ถูกแบนเยอะ)
  var fallback = 0;
  while (prediction.length < 5 && fallback < 100) {
    var r = Math.floor(Math.random() * 10);
    if (prediction.indexOf(r) === -1 && banned.indexOf(r) === -1) {
      prediction.push(r);
    }
    fallback++;
  }

  // สร้าง combos (2 ตัว, 3 ตัว)
  var combosTwo = generateCombos_(prediction, 2);
  var combosThree = generateCombos_(prediction.slice(0, 5), 3); // ใช้ 5 ตัวแรกเพื่อไม่ให้เยอะเกิน

  // บันทึกประวัติ
  saveHistory_(username, lotteryType, mode, input, prediction.join('-'), combosTwo.join(', '), combosThree.join(', '));

  return jsonResponse_({
    status: 'success',
    prediction: prediction
  });
}


// ================================================================
//  🧪 ENGINE — 75 สูตรวิเคราะห์
// ================================================================

/**
 * รัน 75 สูตร และคืนคะแนนเลข 0-9
 * @param {Number[]} d - อาร์เรย์ตัวเลข 6 หลัก
 * @param {String} mode - โหมด A, B, หรือ C
 * @returns {Number[]} คะแนนเลข 0-9
 */
function runAllFormulas_(d, mode) {
  var scores = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

  // ──────────────────────────────────────
  //  กลุ่ม 1: สูตรบวกคงที่ (20 สูตร)
  // ──────────────────────────────────────
  var addResults = [
    (d[0] + d[1]) % 10,                                       // สูตร 1
    (d[1] + d[2]) % 10,                                       // สูตร 2
    (d[2] + d[3]) % 10,                                       // สูตร 3
    (d[3] + d[4]) % 10,                                       // สูตร 4
    (d[4] + d[5]) % 10,                                       // สูตร 5
    (d[0] + d[5]) % 10,                                       // สูตร 6
    (d[0] + d[2]) % 10,                                       // สูตร 7
    (d[1] + d[3]) % 10,                                       // สูตร 8
    (d[2] + d[4]) % 10,                                       // สูตร 9
    (d[3] + d[5]) % 10,                                       // สูตร 10
    (d[0] + d[1] + d[2]) % 10,                                // สูตร 11
    (d[1] + d[2] + d[3]) % 10,                                // สูตร 12
    (d[2] + d[3] + d[4]) % 10,                                // สูตร 13
    (d[3] + d[4] + d[5]) % 10,                                // สูตร 14
    (d[0] + d[3] + d[5]) % 10,                                // สูตร 15
    (d[0] + d[2] + d[4]) % 10,                                // สูตร 16
    (d[1] + d[3] + d[5]) % 10,                                // สูตร 17
    (d[0] + d[1] + d[5]) % 10,                                // สูตร 18
    (d[0] + d[4] + d[5]) % 10,                                // สูตร 19
    (d[0] + d[1] + d[2] + d[3] + d[4] + d[5]) % 10           // สูตร 20
  ];

  // ──────────────────────────────────────
  //  กลุ่ม 2: สูตรคูณทะลวง (15 สูตร)
  // ──────────────────────────────────────
  var mulResults = [
    (d[0] * d[1]) % 10,                                       // สูตร 21
    (d[1] * d[2]) % 10,                                       // สูตร 22
    (d[2] * d[3]) % 10,                                       // สูตร 23
    (d[3] * d[4]) % 10,                                       // สูตร 24
    (d[4] * d[5]) % 10,                                       // สูตร 25
    (d[0] * d[5]) % 10,                                       // สูตร 26
    (d[0] * d[2] + d[1]) % 10,                                // สูตร 27
    (d[1] * d[3] + d[2]) % 10,                                // สูตร 28
    (d[2] * d[4] + d[3]) % 10,                                // สูตร 29
    (d[3] * d[5] + d[4]) % 10,                                // สูตร 30
    (d[0] * d[1] + d[2] * d[3]) % 10,                         // สูตร 31
    (d[1] * d[2] + d[3] * d[4]) % 10,                         // สูตร 32
    (d[2] * d[3] + d[4] * d[5]) % 10,                         // สูตร 33
    (d[0] * d[3] + d[1] * d[5]) % 10,                         // สูตร 34
    ((d[0] * d[1] * d[2]) + (d[3] * d[4] * d[5])) % 10       // สูตร 35
  ];

  // ──────────────────────────────────────
  //  กลุ่ม 3: สูตรบวกไขว้ (25 สูตร)
  // ──────────────────────────────────────
  var crossResults = [
    Math.abs(d[0] - d[1]) % 10,                               // สูตร 36
    Math.abs(d[1] - d[2]) % 10,                               // สูตร 37
    Math.abs(d[2] - d[3]) % 10,                               // สูตร 38
    Math.abs(d[3] - d[4]) % 10,                               // สูตร 39
    Math.abs(d[4] - d[5]) % 10,                               // สูตร 40
    Math.abs(d[0] - d[5]) % 10,                               // สูตร 41
    (d[0] + d[5] + Math.abs(d[2] - d[3])) % 10,              // สูตร 42
    (d[1] + d[4] + Math.abs(d[0] - d[5])) % 10,              // สูตร 43
    (d[0] + d[2] + d[4]) % 10,                                // สูตร 44
    (d[1] + d[3] + d[5]) % 10,                                // สูตร 45
    ((d[0] + d[1]) * (d[4] + d[5])) % 10,                    // สูตร 46
    ((d[0] + d[2]) * (d[3] + d[5])) % 10,                    // สูตร 47
    (d[0] * 2 + d[5]) % 10,                                   // สูตร 48
    (d[1] * 2 + d[4]) % 10,                                   // สูตร 49
    (d[2] * 2 + d[3]) % 10,                                   // สูตร 50
    (d[5] * 2 + d[0]) % 10,                                   // สูตร 51
    (d[0] + d[1] + d[2] + d[3]) % 10,                         // สูตร 52
    (d[1] + d[2] + d[3] + d[4]) % 10,                         // สูตร 53
    (d[2] + d[3] + d[4] + d[5]) % 10,                         // สูตร 54
    Math.abs((d[0] + d[1]) - (d[4] + d[5])) % 10,            // สูตร 55
    Math.abs((d[0] + d[2]) - (d[3] + d[5])) % 10,            // สูตร 56
    (d[0] + d[1] * 2 + d[2]) % 10,                            // สูตร 57
    (d[3] + d[4] * 2 + d[5]) % 10,                            // สูตร 58
    ((d[0] + d[5]) * 2 + d[2]) % 10,                          // สูตร 59
    ((d[1] + d[4]) * 2 + d[3]) % 10                           // สูตร 60
  ];

  // ──────────────────────────────────────
  //  กลุ่ม 4: สูตรเลขเงา (15 สูตร)
  // ──────────────────────────────────────
  var shadowResults = [
    (10 - d[0]) % 10,                                          // สูตร 61
    (10 - d[1]) % 10,                                          // สูตร 62
    (10 - d[2]) % 10,                                          // สูตร 63
    (10 - d[3]) % 10,                                          // สูตร 64
    (10 - d[4]) % 10,                                          // สูตร 65
    (10 - d[5]) % 10,                                          // สูตร 66
    (d[0] + 5) % 10,                                           // สูตร 67
    (d[1] + 5) % 10,                                           // สูตร 68
    (d[2] + 5) % 10,                                           // สูตร 69
    (d[3] + 5) % 10,                                           // สูตร 70
    (d[4] + 5) % 10,                                           // สูตร 71
    (d[5] + 5) % 10,                                           // สูตร 72
    (d[0] + d[5] + 5) % 10,                                   // สูตร 73
    (d[1] + d[4] + 5) % 10,                                   // สูตร 74
    (d[2] + d[3] + 5) % 10                                    // สูตร 75
  ];

  // รวมผลทั้ง 75 สูตร
  var allResults = addResults.concat(mulResults).concat(crossResults).concat(shadowResults);

  // ──────────────────────────────────────
  //  ให้คะแนนตามโหมด Engine
  // ──────────────────────────────────────

  if (mode === 'A') {
    // ⚡ โหมด A: ความถี่สูงสุด
    // นับจำนวนครั้งที่แต่ละเลขออกจาก 75 สูตร
    for (var i = 0; i < allResults.length; i++) {
      scores[allResults[i]] += 1;
    }
  }
  else if (mode === 'B') {
    // ⚡ โหมด B: ลำดับคะแนน
    // สูตรท้ายๆ มีน้ำหนักมากกว่า (ผ่านการกลั่นกรองมากขึ้น)
    for (var i = 0; i < allResults.length; i++) {
      var weight = 1.0 + (i / allResults.length);  // 1.0 → ~2.0
      scores[allResults[i]] += weight;
    }
  }
  else {
    // ⚡ โหมด C: คะแนนดิบ
    // ให้น้ำหนักต่างกันตามกลุ่มสูตร
    var groupWeights = { add: 1.5, mul: 2.0, cross: 2.5, shadow: 1.0 };
    for (var i = 0; i < allResults.length; i++) {
      var w;
      if (i < 20) w = groupWeights.add;
      else if (i < 35) w = groupWeights.mul;
      else if (i < 60) w = groupWeights.cross;
      else w = groupWeights.shadow;
      scores[allResults[i]] += w;
    }
  }

  // ──────────────────────────────────────
  //  โบนัสกำลังวัน
  // ──────────────────────────────────────
  var dayOfWeek = new Date().getDay(); // 0=อาทิตย์ ... 6=เสาร์

  // เลขประจำวัน (ตามตำราโหราศาสตร์ไทย)
  // อา=1, จ=2, อ=3, พ=4, พฤ=5, ศ=6, ส=7
  var dayNumbers = [
    [1, 6],  // อาทิตย์
    [2, 7],  // จันทร์
    [3, 8],  // อังคาร
    [4, 9],  // พุธ
    [5, 0],  // พฤหัสบดี
    [6, 1],  // ศุกร์
    [7, 2]   // เสาร์
  ];

  scores[dayNumbers[dayOfWeek][0]] += 2.0;
  scores[dayNumbers[dayOfWeek][1]] += 1.5;

  // โบนัสเลขจากผลรวม input
  var inputSum = 0;
  for (var i = 0; i < d.length; i++) inputSum += d[i];
  scores[inputSum % 10] += 1.0;
  scores[(inputSum * 3) % 10] += 0.5;

  return scores;
}


// ================================================================
//  🔢 สร้างชุดตัวเลข 2/3 ตัว (Combinations)
// ================================================================

/**
 * สร้าง combinations จาก array
 */
function generateCombos_(arr, size) {
  var result = [];

  function combine(prefix, remaining) {
    for (var i = 0; i < remaining.length; i++) {
      var newPrefix = prefix + String(remaining[i]);
      if (newPrefix.length === size) {
        result.push(newPrefix);
      } else {
        combine(newPrefix, remaining.slice(i + 1));
      }
    }
  }

  combine('', arr);
  return result;
}


// ================================================================
//  📝 ระบบประวัติการวิเคราะห์
// ================================================================

/**
 * บันทึกประวัติลง Sheet "History"
 */
function saveHistory_(username, type, mode, input, result, combosTwo, combosThree) {
  if (!username) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

  sheet.appendRow([username, type, mode, input, result, now, combosTwo, combosThree]);
}

/**
 * ดึงประวัติของ user (ล่าสุด 20 รายการ)
 */
function getHistory_(username) {
  if (!username) return jsonResponse_([]);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  var data = sheet.getDataRange().getValues();
  var results = [];

  // วนจากล่างขึ้นบน เพื่อให้ได้รายการล่าสุดก่อน
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(username)) {
      results.push({
        type:        data[i][1],
        mode:        data[i][2],
        input:       String(data[i][3]),
        result:      String(data[i][4]),
        time:        String(data[i][5]),
        combosTwo:   String(data[i][6] || ''),
        combosThree: String(data[i][7] || '')
      });
    }
    if (results.length >= 20) break;
  }

  return jsonResponse_(results);
}


// ================================================================
//  🧪 ทดสอบ — รันได้ใน Script Editor
// ================================================================

/**
 * ทดสอบระบบวิเคราะห์
 */
function testAnalyze() {
  var testForm = {
    lotteryType: 'รัฐบาลไทย',
    inputNumbers: '835127',
    engineMode: 'A',
    numStates: { '0': 'normal', '1': 'normal', '2': 'normal', '3': 'normal', '4': 'normal', '5': 'normal', '6': 'normal', '7': 'normal', '8': 'normal', '9': 'normal' }
  };

  var result = analyze_(testForm, 'testuser');
  Logger.log(result.getContent());
}

/**
 * ทดสอบระบบสมัครสมาชิก
 */
function testRegister() {
  var result = register_({ username: 'demo', password: '1234' });
  Logger.log(result.getContent());
}

/**
 * ทดสอบระบบเข้าสู่ระบบ
 */
function testLogin() {
  var result = login_({ username: 'demo', password: '1234' });
  Logger.log(result.getContent());
}
