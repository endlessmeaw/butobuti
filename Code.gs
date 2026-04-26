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

  // สร้าง Sheet "lottothai" (เก็บข้อมูลหวยที่ดึงมา)
  let lottoSheet = ss.getSheetByName('lottothai');
  if (!lottoSheet) {
    lottoSheet = ss.insertSheet('lottothai');
    lottoSheet.appendRow(['date', 'first_prize', 'last_two', 'updated_at']);
    lottoSheet.getRange('1:1').setFontWeight('bold').setBackground('#f59e0b').setFontColor('#ffffff');
    lottoSheet.setFrozenRows(1);
  }

  Logger.log('✅ Setup เสร็จสมบูรณ์! สร้าง Sheet "Users", "History" และ "lottothai" แล้ว');
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
  //  รัน 75 สูตร + โบนัสเสริม → คะแนนเลข 0-9
  //  ส่ง lotteryType เพื่อปรับโบนัสวันงวดตามประเภทหวย
  // ================================================
  const scores = runAllFormulas_(digits, mode, lotteryType);

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

  // สร้าง combos (2 ตัว, 3 ตัว) จาก 5 ตัวแรก
  var top5 = prediction.slice(0, 5);
  var combosTwo   = generateCombos_(top5, 2);
  var combosThree = generateCombos_(top5, 3);

  // บันทึกประวัติ
  saveHistory_(username, lotteryType, mode, input, prediction.join('-'), combosTwo.join(', '), combosThree.join(', '));

  return jsonResponse_({
    status: 'success',
    prediction: prediction,              // [2, 5, 3, 1, 7, 0] (ทั้งหมด 6 ตัว)
    topTwo: prediction.slice(0, 2),      // [2, 5] (เด่น 2 ตัว)
    topFive: prediction.slice(0, 5),     // [2, 5, 3, 1, 7] (วิน 5 ตัว)
    combosTwo: combosTwo,                // ['25', '23', '21', ...] (เจาะ 2 ตัว)
    combosThree: combosThree             // ['253', '251', ...] (เจาะ 3 ตัว)
  });
}


// ================================================================
//  🧪 ENGINE — 75 สูตรวิเคราะห์
// ================================================================

/**
 * รัน 75 สูตร และคืนคะแนนเลข 0-9
 * @param {Number[]} d           - อาร์เรย์ตัวเลข 6 หลัก
 * @param {String}   mode        - โหมด A, B, หรือ C
 * @param {String}   lotteryType - ประเภทหวย เพื่อคำนวณวันงวดถัดไป
 * @returns {Number[]} คะแนนเลข 0-9
 */
function runAllFormulas_(d, mode, lotteryType) {
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

  // ══════════════════════════════════════
  //  🔥 Hot Pair Analysis (ใหม่!)
  //  อิงจากตารางความถี่คู่เลข 2 ตัว (00-99) ย้อนหลัง
  //  เลขสีเข้มในตาราง = ออกบ่อย → เลขในคู่นั้นได้โบนัส
  // ══════════════════════════════════════

  // ดึงข้อมูลความถี่ของตัวเลขเดี่ยว และคู่เลขเด่น จาก Sheet lottothai
  // หากยังไม่มีข้อมูล ระบบจะใช้ค่า Default แทน
  var hotStats = calculateHotStatsFromSheet_();
  var hotDigitBase = hotStats.digitBase;
  var hotPairScores = hotStats.pairScores;


  // 1) โบนัสฐาน hot digit (ทุกเลข 0-9)
  for (var i = 0; i <= 9; i++) {
    scores[i] += hotDigitBase[i] * 0.3;
  }

  // 2) หาเลขที่ไม่ซ้ำใน input
  var inputDigitSet = [];
  var seenDigit = {};
  for (var i = 0; i < d.length; i++) {
    if (!seenDigit[d[i]]) { seenDigit[d[i]] = true; inputDigitSet.push(d[i]); }
  }

  // 3) โบนัสเพิ่มถ้าเลขดังอยู่ใน input ด้วย (Intersection Boost)
  for (var i = 0; i < inputDigitSet.length; i++) {
    scores[inputDigitSet[i]] += hotDigitBase[inputDigitSet[i]] * 0.25;
  }

  // 4) โบนัส Hot Pair จากคู่ตัวเลขใน input
  //    ถ้า input มีทั้งเลข A และ B และ AB เป็น hot pair → A,B ได้โบนัส
  for (var a = 0; a < inputDigitSet.length; a++) {
    for (var b = 0; b < inputDigitSet.length; b++) {
      if (a === b) continue;
      var pair = String(inputDigitSet[a]) + String(inputDigitSet[b]);
      var ps = hotPairScores[pair] || 0;
      if (ps > 0) {
        scores[inputDigitSet[a]] += ps * 0.35;
        scores[inputDigitSet[b]] += ps * 0.35;
      }
    }
  }

  // 5) โบนัสเลขนอก input ที่สร้าง hot pair กับ input digits
  //    เช่น input มี 5 และ 57 hot → เลข 7 ได้โบนัสแม้ไม่อยู่ใน input
  for (var x = 0; x <= 9; x++) {
    var externalHotTotal = 0;
    for (var a = 0; a < inputDigitSet.length; a++) {
      var p1 = String(x) + String(inputDigitSet[a]);
      var p2 = String(inputDigitSet[a]) + String(x);
      var best = Math.max(hotPairScores[p1] || 0, hotPairScores[p2] || 0);
      externalHotTotal += best;
    }
    // ใช้ค่าเฉลี่ยคูณ weight สูง เพื่อให้เลขที่มีคู่ดังหลายคู่ได้โบนัสสะสม
    scores[x] += externalHotTotal * 0.35;
  }

  // ──────────────────────────────────────
  //  โบนัสเลขที่ปรากฏใน input (Digit Presence Bonus)
  // ──────────────────────────────────────
  var digitCount = [0,0,0,0,0,0,0,0,0,0];
  for (var i = 0; i < d.length; i++) { digitCount[d[i]] += 1; }
  for (var i = 0; i <= 9; i++) {
    if (digitCount[i] >= 2) scores[i] += 1.5;   // ลดจาก 2.0 → 1.5
    else if (digitCount[i] === 1) scores[i] += 0.7; // ลดจาก 1.0 → 0.7
  }

  // ──────────────────────────────────────
  //  โบนัสเลขสะท้อน (Mirror / Positional)
  // ──────────────────────────────────────
  scores[(d[4] + d[5]) % 10] += 1.0;           // เลขท้าย 2 ตัวรวม
  scores[(d[0] + d[1] + d[2]) % 10] += 0.8;    // หัว 3 ตัวรวม
  scores[(d[3] + d[4] + d[5]) % 10] += 0.8;    // ท้าย 3 ตัวรวม

  // ──────────────────────────────────────
  //  โบนัสกำลังวัน — คำนวณตามงวดถัดไปของแต่ละประเภทหวย
  // ──────────────────────────────────────
  var now = new Date();
  var dayOfMonth = now.getDate();
  var nextDrawDate;

  if (lotteryType === 'รัฐบาลไทย' || lotteryType === 'หวยออมสิน') {
    // งวดออก 1 และ 16 ของทุกเดือน
    if (dayOfMonth < 16) {
      nextDrawDate = new Date(now.getFullYear(), now.getMonth(), 16);
    } else {
      nextDrawDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    }
  } else if (lotteryType === 'ลาวพัฒนา') {
    // ลาวพัฒนาออกทุกวันจันทร์, พุธ, ศุกร์
    // หาวันถัดไปที่เป็น จ,พ,ศ
    var daysToNext = 1;
    for (var trial = 1; trial <= 7; trial++) {
      var candidate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + trial);
      var cDay = candidate.getDay(); // 1=จ, 3=พ, 5=ศ
      if (cDay === 1 || cDay === 3 || cDay === 5) { nextDrawDate = candidate; break; }
    }
    if (!nextDrawDate) nextDrawDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  } else if (lotteryType === 'ฮานอยอาเซียน') {
    // ฮานอยออกทุกวัน (วันถัดไปเสมอ)
    nextDrawDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  } else {
    // default → พรุ่งนี้
    nextDrawDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  }

  var drawDayOfWeek = nextDrawDate.getDay();

  // เลขประจำวัน (ตามตำราโหราศาสตร์ไทย)
  var dayNumbers = [
    [1, 6],  // อาทิตย์
    [2, 7],  // จันทร์
    [3, 8],  // อังคาร
    [4, 9],  // พุธ
    [5, 0],  // พฤหัสบดี
    [6, 1],  // ศุกร์
    [7, 2]   // เสาร์
  ];

  scores[dayNumbers[drawDayOfWeek][0] % 10] += 2.0;
  scores[dayNumbers[drawDayOfWeek][1] % 10] += 1.5;

  // ──────────────────────────────────────
  //  โบนัสผลรวม input
  // ──────────────────────────────────────
  var inputSum = 0;
  for (var i = 0; i < d.length; i++) inputSum += d[i];
  scores[inputSum % 10] += 1.5;
  scores[(inputSum % 10 + 5) % 10] += 0.5;  // เลขคู่กับผลรวม

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
//  🌐 ระบบ Scraping (ดึงข้อมูลหวยอัตโนมัติ)
// ================================================================

/**
 * ดึงข้อมูลรางวัลที่ 1 งวดล่าสุดจากเว็บ Sanook และบันทึกลง Sheet "lottothai"
 * เพื่อเป็นแคช (Cache) จะได้ไม่ต้องโหลดจากเว็บทุกครั้ง
 * @returns {Object} ข้อมูลที่ดึงมา {date, firstPrize, lastTwo} หรือ null
 */
function getLatestLottoData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('lottothai');
  
  if (!sheet) {
    setupSheets(); // เผื่อยังไม่ได้สร้าง Sheet
    sheet = ss.getSheetByName('lottothai');
  }

  // 1. ลองดึงข้อมูลจากแถวบนสุด (Row 2) ใน Sheet ดูก่อน (เพราะข้อมูลใหม่จะอยู่บนสุดเสมอ)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var data = sheet.getRange(2, 1, 1, 4).getValues()[0];
    var savedDate = data[0];
    var savedFirstPrize = String(data[1]).padStart(6, '0');
    var savedLastTwo = String(data[2]).padStart(2, '0');
    
    // ตรวจสอบว่าอัปเดตล่าสุดเมื่อไหร่ ถ้าเพิ่งอัปเดตวันนี้ ไม่ต้องไปดึงใหม่
    var updatedAt = new Date(data[3]);
    var now = new Date();
    var diffHours = (now - updatedAt) / (1000 * 60 * 60);
    
    // ถ้าแคชยังใหม่อยู่ (ดึงมาไม่เกิน 12 ชั่วโมง) ให้ใช้ค่าจาก Sheet ได้เลย
    if (diffHours < 12) {
      Logger.log("ใช้ข้อมูลจาก Sheet (Cache): งวด " + savedDate);
      return {
        date: savedDate,
        firstPrize: savedFirstPrize,
        lastTwo: savedLastTwo,
        fromCache: true
      };
    }
  }

  // 2. ถ้าไม่มีข้อมูลใน Sheet หรือข้อมูลเก่าเกิน 12 ชั่วโมง ให้ดึงจากเว็บ
  try {
    var url = "https://news.sanook.com/lotto/archive/";
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText();
    
    // Regex จับคู่วันที่ และ เลข 6 หลักรางวัลที่ 1 (ตัวแรกสุดที่หาเจอและเป็นตัวเลข 6 หลัก)
    // ข้ามงวดที่ยังไม่ออก (ที่เป็น xxxxxx)
    var regex = /ตรวจหวย\s+(\d{1,2}\s+[^\s]+\s+\d{4})[\s\S]{0,200}?รางวัลที่ 1[\s\S]{0,100}?>\s*(\d{6})\s*</g;
    var match = regex.exec(html);
    
    if (match && match[1] && match[2]) {
      var drawDate = match[1];         // เช่น "16 เมษายน 2569"
      var firstPrize = match[2];       // เช่น "309612"
      var lastTwo = firstPrize.slice(-2);
      
      // เช็คว่าตรงกับข้อมูลแถวที่ 2 ใน Sheet ไหม ถ้าไม่ตรงค่อยเพิ่มแถวใหม่
      var isNewDraw = true;
      if (lastRow > 1 && savedDate === drawDate) {
        isNewDraw = false;
        // อัปเดตเวลาอัปเดตแถวที่ 2 (ข้อมูลเดิม)
        sheet.getRange(2, 4).setValue(new Date()); 
      }
      
      if (isNewDraw) {
        var nowTimestamp = new Date();
        sheet.insertRowAfter(1); // แทรกแถวใหม่ต่อจากหัวตาราง
        sheet.getRange(2, 1, 1, 4).setValues([[drawDate, firstPrize, lastTwo, nowTimestamp]]);
        Logger.log("บันทึกข้อมูลใหม่ลง Sheet: งวด " + drawDate + " รางวัลที่ 1: " + firstPrize);
      }
      
      return {
        date: drawDate,
        firstPrize: firstPrize,
        lastTwo: lastTwo,
        fromCache: false
      };
    } else {
      Logger.log("ไม่พบข้อมูลจากการ Scrape (อาจเป็นรูปแบบเว็บเปลี่ยน)");
    }
  } catch (error) {
    Logger.log("เกิดข้อผิดพลาดในการ Scraping: " + error.message);
  }
  
  // 3. ถ้าเกิด Error ในการ Scraping ให้ส่งข้อมูลแถวบนสุด (Row 2) จาก Sheet กลับไปแทน (ถ้ามี)
  if (lastRow > 1) {
    var fallbackData = sheet.getRange(2, 1, 1, 3).getValues()[0];
    return {
      date: fallbackData[0],
      firstPrize: String(fallbackData[1]).padStart(6, '0'),
      lastTwo: String(fallbackData[2]).padStart(2, '0'),
      fromCache: true
    };
  }
  
  return null;
}

/**
 * ดึงข้อมูลหวยย้อนหลังหลายหน้า (เช่น 30 หน้า = ประมาณ 600 งวด / 25 ปี)
 * @param {Number} maxPages - จำนวนหน้าที่ต้องการดึง (1 หน้ามีประมาณ 20 งวด)
 */
function scrapeHistoricalLottoData_(maxPages) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('lottothai');
  
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('lottothai');
  }
  
  var max = maxPages || 30; // ค่าเริ่มต้นดึง 30 หน้า
  var newRows = [];
  
  // 1. ดึงวันที่ที่มีอยู่แล้วใน Sheet มาเก็บไว้ เพื่อกันการบันทึกซ้ำ
  var existingDates = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var savedData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < savedData.length; i++) {
      existingDates[savedData[i][0]] = true;
    }
  }

  Logger.log("เริ่มดึงข้อมูลย้อนหลังจำนวน " + max + " หน้า...");

  // 2. ลูปดึงทีละหน้า
  for (var page = 1; page <= max; page++) {
    try {
      var url = (page === 1) 
        ? "https://news.sanook.com/lotto/archive/" 
        : "https://news.sanook.com/lotto/archive/page/" + page + "/";
        
      var response = UrlFetchApp.fetch(url);
      var html = response.getContentText();
      
      var regex = /ตรวจหวย\s+(\d{1,2}\s+[^\s]+\s+\d{4})[\s\S]{0,200}?รางวัลที่ 1[\s\S]{0,100}?>\s*(\d{6})\s*</g;
      var match;
      var foundInPage = 0;
      
      while ((match = regex.exec(html)) !== null) {
        var drawDate = match[1];
        var firstPrize = match[2];
        var lastTwo = firstPrize.slice(-2);
        
        // ถ้ายังไม่มีใน Sheet ให้เก็บไว้เตรียมบันทึก
        if (!existingDates[drawDate]) {
          var nowTimestamp = new Date();
          newRows.push([drawDate, firstPrize, lastTwo, nowTimestamp]);
          existingDates[drawDate] = true; // มาร์คไว้ว่าเพิ่งเจอ เพื่อไม่ให้ซ้ำในรอบถัดไป
          foundInPage++;
        }
      }
      
      Logger.log("หน้า " + page + ": ดึงข้อมูลใหม่ได้ " + foundInPage + " งวด");
      
    } catch (e) {
      Logger.log("หน้า " + page + " Error: " + e.message);
      break; // ถ้าพัง (เช่น ทะลุหน้าสุดท้าย) ให้หยุดลูป
    }
  }
  
  // 3. แทรกข้อมูลใหม่ไว้ด้านบนสุด (ต่อจากหัวตาราง)
  if (newRows.length > 0) {
    sheet.insertRowsAfter(1, newRows.length);
    sheet.getRange(2, 1, newRows.length, 4).setValues(newRows);
    Logger.log("✅ บันทึกข้อมูลย้อนหลังสำเร็จทั้งหมด " + newRows.length + " งวด!");
  } else {
    Logger.log("✅ ไม่มีข้อมูลงวดใหม่ให้บันทึก (ข้อมูลอัปเดตล่าสุดอยู่แล้ว)");
  }
}

/**
 * คำนวณความถี่จากข้อมูลจริงใน Sheet "lottothai"
 * นำเลขท้าย 2 ตัวของรางวัลที่ 1 มานับความถี่เพื่อใช้แทน Hot Pair แบบเดิม
 */
function calculateHotStatsFromSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss ? ss.getSheetByName('lottothai') : null;
  
  // ค่า Default เดิม กรณีไม่มี Sheet หรือยังไม่เคยดึงข้อมูล
  var defaultStats = {
    digitBase: [5, 9, 9, 7, 4, 8, 4, 7, 4, 4],
    pairScores: {
      '21':10, '12':8, '10':9, '01':8, '20':9, '02':8,
      '53':9, '35':8, '57':9, '75':8, '43':7, '34':6,
      '51':7, '15':6, '41':5, '14':4, '42':6, '24':5,
      '13':7, '31':7, '23':7, '32':6, '17':8, '71':7,
      '47':7, '74':6, '25':6, '52':6, '86':6, '68':5,
      '87':6, '78':5, '91':5, '19':5, '76':7, '67':6,
      '97':5, '79':5, '46':4, '64':4
    }
  };

  if (!sheet) return defaultStats;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return defaultStats; 
  
  // ดึงข้อมูลเลขท้าย 2 ตัวมาทั้งหมด (คอลัมน์ที่ 3)
  var data = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  
  var pairCounts = {};
  var digitCounts = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  var maxPair = 0;
  var maxDigit = 0;
  
  for (var i = 0; i < data.length; i++) {
    var pair = String(data[i][0]).padStart(2, '0');
    if (pair.length === 2) {
      pairCounts[pair] = (pairCounts[pair] || 0) + 1;
      if (pairCounts[pair] > maxPair) maxPair = pairCounts[pair];
      
      var d1 = parseInt(pair[0], 10);
      var d2 = parseInt(pair[1], 10);
      if (!isNaN(d1)) {
        digitCounts[d1]++;
        if (digitCounts[d1] > maxDigit) maxDigit = digitCounts[d1];
      }
      if (!isNaN(d2)) {
        digitCounts[d2]++;
        if (digitCounts[d2] > maxDigit) maxDigit = digitCounts[d2];
      }
    }
  }
  
  // แปลงให้เป็นสเกล 0-10 (เพื่อให้คู่ที่ออกบ่อยสุดได้ 10 คะแนน เทียบเท่าของเดิม)
  var normalizedDigitBase = [];
  for (var d = 0; d <= 9; d++) {
    normalizedDigitBase[d] = maxDigit > 0 ? Math.round((digitCounts[d] / maxDigit) * 10) : 0;
  }
  
  var normalizedPairScores = {};
  for (var p in pairCounts) {
    if (maxPair > 0) {
      var score = Math.round((pairCounts[p] / maxPair) * 10);
      if (score >= 4) { // กรองเอาเฉพาะคู่ที่มีคะแนนระดับ 4 ขึ้นไป
        normalizedPairScores[p] = score;
      }
    }
  }
  
  if (maxPair > 0) {
    return {
      digitBase: normalizedDigitBase,
      pairScores: normalizedPairScores
    };
  }
  
  return defaultStats;
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

/**
 * ทดสอบดึงข้อมูลสลากกินแบ่งรัฐบาลล่าสุดและบันทึกลง lottothai
 */
function testScrapingLotto() {
  var result = getLatestLottoData_();
  if (result) {
    Logger.log("🎉 ดึงข้อมูลสำเร็จ!");
    Logger.log("วันที่: " + result.date);
    Logger.log("รางวัลที่ 1: " + result.firstPrize);
    Logger.log("เลขท้าย 2 ตัว: " + result.lastTwo);
    Logger.log("ใช้ข้อมูลจาก Cache ใน Sheet หรือไม่?: " + (result.fromCache ? "ใช่" : "ไม่ใช่ (ดึงใหม่จากเว็บ)"));
  } else {
    Logger.log("❌ ไม่สามารถดึงข้อมูลได้");
  }
}

/**
 * ทดสอบดึงข้อมูลหวยย้อนหลังรวดเดียวหลายสิบปี!
 * ปรับตัวเลข 30 เป็นจำนวนหน้าที่ต้องการ (1 หน้ามีประมาณ 20 งวด)
 * 30 หน้า = ประมาณ 600 งวด (ย้อนหลังประมาณ 25 ปี)
 */
function testScrapeHistory() {
  // ใส่ตัวเลขจำนวนหน้า เช่น 30 หน้า
  // หากต้องการดึงให้สุดเว็บ อาจจะใส่ไปเลย 100 หน้า แต่ระวังอาจรันนานเกิน 6 นาที
  scrapeHistoricalLottoData_(30);
}
