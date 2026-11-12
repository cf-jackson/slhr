const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const functions = [
  'getAvailableShifts', 'normalizeImportedShift', 'getShiftForDay',
  'resetImportedShiftOverrides', 'normalizeHeader', 'pick', 'excelDateToDate',
  'formatDate', 'monthKey', 'timeText', 'durationToHours', 'parseClockMinutes',
  'estimateHours', 'calculateShiftTotalHours', 'getPunchIssue',
  'normalizePersonText', 'findPersonShiftName', 'parsePersonShiftSheet',
  'overtimeMultipliedHours', 'getOvertimeLimit', 'redecodeBig5Text', 'redecodeWorkbookBig5', 'comparePeople', 'personShiftLabel',
  'getPersonUsualShiftName', 'joinNotes', 'hasAnyPunch', 'getLateEarlyMinutes',
  'summarizeAnomalies', 'anomalyItems', 'getDayType', 'restHolidayOvertimeHours',
  'applyOvertimeRule', 'parseOvertime', 'normalizeRows', 'datesInMonth',
  'getReport', 'shiftHours', 'calculateShiftOvertime', 'formatHours',
  'formatDayTotalHours', 'getReportDayStatus', 'isCalendarRestDay',
  'withHolidayFill', 'thinBorder', 'formatClockMinutes', 'getInspectionStartTime',
  'getInspectionStatus', 'getAdjustedInspectionEnd', 'loadWorkbook', 'readWorkbookFile'
];

function harness() {
  const callbacks = {};
  const exports = [];
  const defaults = [
    { name: 'A', start: '08:30', end: '17:30', lunchMinutes: 60 },
    { name: 'B', start: '08:30', end: '17:00', lunchMinutes: 60 },
    { name: 'C', start: '08:20', end: '17:30', lunchMinutes: 70 },
    { name: 'P', start: '', end: '', lunchMinutes: 60 }
  ];
  const context = vm.createContext({
    Date, Blob, console, TextDecoder,
    defaultShifts: defaults, shifts: defaults.map((s) => ({ ...s })),
    shiftOverrides: {}, personShifts: [], DEFAULT_OVERTIME_LIMIT: 54, REST_DAY_SCHEDULE: { start: '08:00', end: '16:30', lunchMinutes: 30, lunchStart: '12:00', lunchEnd: '12:30' }, NO_SHOW_LABEL: '未出勤（請假／曠職待確認）', rawRows: [], selectedEmployeeKey: '',
    monthSelect: { value: '2026-09' }, overtimeLimitInput: { value: '54' },
    floorOvertimeCheckbox: { checked: true }, weekdays: ['日', '一', '二', '三', '四', '五', '六'],
    getHolidayForDate: () => null, updateMonthOptions() {}, render() {},
    localStorage: { setItem() {} },
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL() {} },
    document: { body: { appendChild() {}, removeChild() {} }, createElement: () => ({ click() {} }) },
    exportButton: { addEventListener: (_, cb) => { callbacks.B = cb; } },
    complianceButton: { addEventListener: (_, cb) => { callbacks.A = cb; } },
    XLSX: {
      utils: {
        sheet_to_json: (sheet, options) => options.header === 1 ? [Object.keys(sheet[0])] : sheet,
        book_new: () => ({ sheets: [] }),
        aoa_to_sheet: (rows) => {
          const sheet = {};
          rows.forEach((row, r) => row.forEach((v, c) => { sheet[`${r}:${c}`] = { v }; }));
          return sheet;
        },
        decode_cell: (address) => {
          const [r, c] = address.split(':').map(Number);
          return { r, c };
        },
        book_append_sheet: (book, sheet) => book.sheets.push(sheet)
      },
      write: (book) => { exports.push(book); return ''; }
    }
  });
  for (const name of functions) {
    const start = html.indexOf(`    function ${name}(`);
    assert.notEqual(start, -1, name);
    const end = html.indexOf('\n    }', start) + '\n    }'.length;
    vm.runInContext(html.slice(start, end), context);
  }
  for (const [startMarker, endMarker] of [
    ['    exportButton.addEventListener("click"', '    function formatClockMinutes'],
    ['    complianceButton.addEventListener("click"', '    function thinBorder']
  ]) {
    const start = html.indexOf(startMarker);
    vm.runInContext(html.slice(start, html.indexOf(endMarker, start)), context);
  }
  return { context, callbacks, exports };
}

function input(day, period, start = '08:30', end = '17:30', extra = {}) {
  const row = { 姓名: '測試員工', 工號: 'T1', 日期: new Date(2026, 8, day), 上班1: start, 下班1: end, ...extra };
  if (period !== undefined) row[' 對應時段 '] = period;
  return row;
}

function load(c, rows) {
  c.loadWorkbook({ SheetNames: ['data'], Sheets: { data: rows } });
  return c.getReport().people[0].days;
}

test('import mapping drives shifts, overtime multiplier and both export shift columns', () => {
  const { context: c, callbacks, exports } = harness();
  // 2026/9/1 週二。A 班 08:30–17:30，19:45 下班 → 加班 2 小時 15 分，無條件捨去為 2。
  // 9/7 是週一且完全沒有「對應時段」欄，確認會退回第一個班別。
  const inputs = [input(1, 'A', '08:30', '19:45'), input(2, ' B '), input(3, '其他'), input(4, ''), input(7)];
  const expected = ['A', 'B', 'A', 'A'];
  c.loadWorkbook({ SheetNames: ['data'], Sheets: { data: inputs } });
  assert.deepEqual(Array.from(c.rawRows, (r) => r.shiftName), ['A', 'B', '', '', '']);
  const person = c.getReport().people[0];
  assert.deepEqual(Array.from(person.days.slice(0, 4), (d) => d.shift.name), expected);
  assert.equal(person.days[6].shift.name, 'A');
  assert.equal(person.days[0].overtimeHours, 2);
  assert.equal(person.days[0].overtimeMultiplied, 2.68);
  assert.equal(person.overtimeTotal, 2);
  assert.equal(person.overtimeMultipliedTotal, 2.68);
  assert.equal(person.days[5].shift.name, 'A');
  callbacks.B();
  callbacks.A();
  // B 表第 1 張是「異常匯總」總表，個人表從第 2 張開始；A 表沒有總表。
  for (const sheet of [exports[0].sheets[1], exports[1].sheets[0]]) {
    assert.equal(sheet['0:2'].v, '考勤號碼：');
    assert.equal(sheet['0:4'].v, '員工序號：T1');
    assert.equal(sheet['1:2'].v, '班別');
    expected.forEach((name, i) => assert.equal(sheet[`${i + 2}:2`].v, name));
  }
  // B 表：新增「加班倍數後」欄，並在每位員工下方有兩列合計（9 月 30 天 → 第 32、33 列）。
  const b = exports[0].sheets[1];
  assert.equal(b['1:12'].v, '加班倍數後');
  assert.equal(b['2:12'].v, '2.68');
  assert.equal(b['32:0'].v, '原加班時數合計');
  assert.equal(b['32:11'].v, '2');
  assert.equal(b['33:0'].v, '加班倍數後合計');
  assert.equal(b['33:12'].v, '2.68');
});

test('re-import replaces affected month overrides but permits later manual changes', () => {
  const { context: c } = harness();
  c.shiftOverrides = {
    'T1|測試員工|2026/9/1': 'B', 'T1|測試員工|2026/9/30': 'B',
    'T1|測試員工|2026/8/1': 'B', 'other|2026/9/1': 'B'
  };
  const book = { SheetNames: ['data'], Sheets: { data: [input(1, 'A')] } };
  c.loadWorkbook(book);
  assert.equal(c.shiftOverrides['T1|測試員工|2026/9/30'], undefined);
  assert.equal(c.shiftOverrides['T1|測試員工|2026/8/1'], 'B');
  assert.equal(c.getReport().people[0].days[0].shift.name, 'A');
  c.shiftOverrides['T1|測試員工|2026/9/1'] = 'B';
  assert.equal(c.getReport().people[0].days[0].shift.name, 'B');
  c.loadWorkbook(book);
  assert.equal(c.getReport().people[0].days[0].shift.name, 'A');
});

test('configured shift names are matched dynamically and built-ins remain selectable', () => {
  const { context: c } = harness();
  c.shifts = [
    { name: '日班', start: '09:00', end: '18:00', lunchMinutes: 60 },
    { name: '晚班', start: '14:00', end: '22:00', lunchMinutes: 0 }
  ];
  const days = load(c, [input(1, '日班'), input(2, '晚班', '14:00', '22:00'), input(3, '未知班')]);
  assert.equal(days[0].shift.start, '09:00');
  assert.equal(days[1].shift.name, '晚班');
  assert.equal(days[1].shift.end, '22:00');
  assert.equal(days[2].shift.name, '日班');
  assert.ok(c.getAvailableShifts().some((s) => s.name === 'C'));
});

test('required hours are shift span minus lunch; short only when below that', () => {
  const { context: c } = harness();
  const [a, b, cShift, p] = c.getAvailableShifts();
  assert.deepEqual([a, b, cShift].map((s) => c.shiftHours(s)), [8, 7.5, 8]);
  assert.equal(c.shiftHours(p), null);
  const days = load(c, [
    input(1, 'A', '08:30', '17:30'),   // 9 小時 − 1 = 8，剛好不算未滿
    input(2, 'A', '08:30', '17:00'),   // 8.5 − 1 = 7.5 < 8 → 未滿
    input(3, 'B', '08:30', '17:00'),   // B 應上 7.5，剛好
    input(4, 'C', '08:20', '17:30'),   // 9h10 − 70 分 = 8，剛好
    input(7, 'C', '08:20', '17:00'),   // 9/7 週一：8h40 − 70 分 = 7.5 < 8 → 未滿
    input(5, 'C', '08:20', '17:00')    // 9/5 週六：同樣 7.5 小時，但休息日不判定未滿
  ]);
  const picked = [0, 1, 2, 3, 6];
  assert.deepEqual(picked.map((i) => days[i].totalHours), [8, 7.5, 7.5, 8, 7.5]);
  assert.deepEqual(picked.map((i) => days[i].short), [false, true, false, false, true]);
  assert.equal(days[4].short, false);
  assert.equal(days[1].note, '未滿 8 小時；早退 30 分');
});

test('flexible P shift shows hours but is never short and never earns overtime', () => {
  const { context: c } = harness();
  const days = load(c, [
    input(1, 'P', '09:00', '12:00'),     // 週二，沒跨午休 → 不扣午休，3 小時
    input(2, 'P', '08:00', '20:00'),     // 很晚下班也不算加班
    input(5, 'P', '08:00', '16:00')      // 週六：一般班別整天計加班，P 不計
  ]);
  assert.equal(days[0].totalHours, 3);
  assert.equal(days[0].short, false);
  assert.equal(days[1].overtimeHours, 0);
  assert.equal(days[4].shift.name, 'P');
  assert.equal(days[4].overtimeHours, 0);
  assert.equal(days[4].overtimeMultiplied, 0);
});

test('lunch is deducted once only when a single punch pair covers noon', () => {
  const { context: c } = harness();
  const shiftA = c.getAvailableShifts()[0];
  assert.equal(c.calculateShiftTotalHours({ start1: '08:30', end1: '17:30' }, shiftA), 8);
  assert.equal(c.calculateShiftTotalHours({ start1: '08:30', end1: '12:00' }, shiftA), 3.5);
  assert.equal(c.calculateShiftTotalHours({ start1: '08:30', end1: '12:30', start2: '13:30', end2: '17:30' }, shiftA), 8);
  assert.equal(c.calculateShiftTotalHours({ start1: '', end1: '', totalHours: 6 }, shiftA), 6);
});

test('single-sided punches are flagged and not counted; other columns are not used to guess', () => {
  const { context: c } = harness();
  const days = load(c, [
    input(1, 'A', '', '23:07', { 出勤時間: '14:37', 合計: '14:37' }),
    input(2, 'A', '08:27', ''),
    input(3, 'A', '', '')
  ]);
  assert.equal(days[0].missingPunch, true);
  assert.equal(days[0].note, '上班未打卡');
  assert.equal(days[0].totalHours, 0);
  assert.equal(days[0].overtimeHours, 0);
  assert.equal(days[0].short, false);
  assert.equal(days[1].note, '下班未打卡');
  assert.equal(days[2].absent, true);
  assert.equal(c.getReportDayStatus(days[0], 54), '上班未打卡');
  assert.equal(c.getReportDayStatus(days[1], 54), '下班未打卡');
});

test('overtime multiplier tiers by day type', () => {
  const { context: c } = harness();
  // 工作日、休息日：前 2 小時 ×1.34，第 3 小時起 ×1.67
  assert.deepEqual([1, 2, 3, 4].map((h) => c.overtimeMultipliedHours(h, 'workday')), [1.34, 2.68, 4.35, 6.02]);
  assert.equal(c.overtimeMultipliedHours(4, 'restDay'), 6.02);
  // 休假日：前 8 小時 ×1，第 9～10 小時 ×1.34，第 11 小時起 ×1.67（超過 12 小時比照 ×1.67）
  assert.deepEqual([8, 9, 10, 11, 12, 13].map((h) => c.overtimeMultipliedHours(h, 'restHoliday')), [8, 9.34, 10.68, 12.35, 14.02, 15.69]);
  assert.equal(c.overtimeMultipliedHours(5, 'typhoon'), 0);
  assert.equal(c.overtimeMultipliedHours(0, 'workday'), 0);
  // 休假日 8 小時內一律計 8 小時，超過照實際
  assert.deepEqual([0, 2, 8, 10.5].map((h) => c.restHolidayOvertimeHours(h)), [0, 8, 8, 10.5]);
  // 整月：週六 8 小時 = 2×1.34 + 6×1.67；週日只上 2 小時也計 8；國定假日（9/2 週三）10 小時 = 8 + 2×1.34
  c.getHolidayForDate = (date) => (date.getDate() === 2 ? { name: '測試假日', type: 'national' } : null);
  load(c, [input(5, 'A', '08:30', '17:30'), input(6, 'A', '08:30', '10:30'), input(2, 'A', '08:30', '19:30')]);
  const person = c.getReport().people[0];
  assert.equal(person.days[4].overtimeHours, 8);
  assert.equal(person.days[4].overtimeMultiplied, 12.7);
  assert.equal(person.days[4].note, '休息日出勤全數計加班');
  assert.equal(person.days[5].totalHours, 2);
  assert.equal(person.days[5].overtimeHours, 8);
  assert.equal(person.days[5].overtimeMultiplied, 8);
  assert.equal(person.days[5].note, '休假日出勤，前 8 小時直接計 8 小時');
  assert.equal(person.days[1].overtimeHours, 10);
  assert.equal(person.days[1].overtimeMultiplied, 10.68);
  assert.equal(person.days[1].note, '測試假日：休假日出勤，前 8 小時直接計 8 小時');
  assert.equal(person.overtimeTotal, 26);                 // 全部計入月上限累計
  assert.equal(person.overtimeMultipliedTotal, 31.38);
});

// 用 2611.xls 11/7（週六）的真實打卡時間；若沒套用表上時段，C 班會從 08:20 起算、扣 70 分，
// 得到 7.02／3.75 小時（捨去後加班 7／3），這裡的斷言就會失敗。
test('Saturday and rest holidays use the table schedule: from 08:00, lunch 12:00–12:30 (30 min)', () => {
  const { context: c, callbacks, exports } = harness();
  c.getHolidayForDate = (date) => (date.getDate() === 2 ? { name: '測試假日', type: 'national' } : null);
  const days = load(c, [
    input(5, 'C', '07:48', '16:31'),   // 週六：08:00～16:31 = 8h31 − 30 分 = 8.02
    input(12, 'C', '07:43', '12:05'),  // 週六半天：08:00～12:05 = 4.08，沒跨 12:00～12:30 不扣
    input(6, 'C', '07:30', '19:00'),   // 週日：08:00～19:00 = 11h − 30 分 = 10.5 → 捨去 10
    input(2, 'A', '08:00', '17:00'),   // 國定假日（週三）：08:00～17:00 = 9h − 30 分 = 8.5 → 捨去 8 → 計 8
    input(7, 'C', '07:48', '16:31')    // 週一平日：仍用 C 班 08:20 起、扣 70 分 = 7.02
  ]);
  assert.equal(c.formatHours(days[4].totalHours), '8.02');
  assert.equal(days[4].overtimeHours, 8);
  assert.equal(days[4].overtimeMultiplied, 12.7);
  assert.equal(c.formatHours(days[11].totalHours), '4.08');
  assert.equal(days[11].overtimeHours, 4);
  assert.equal(days[5].totalHours, 10.5);
  assert.equal(days[5].overtimeHours, 10);
  assert.equal(days[5].overtimeMultiplied, 10.68);
  assert.equal(days[1].totalHours, 8.5);
  assert.equal(days[1].overtimeHours, 8);
  assert.equal(c.formatHours(days[6].totalHours), '7.02');
  assert.equal(days[4].shift.name, 'C');                   // 班別仍是 C
  // P 班不套用表上時段：週六 07:00～12:00 仍是 5 小時、不計加班
  c.shiftOverrides['T1|測試員工|2026/9/19'] = 'P';
  const pDays = load(c, [input(19, 'P', '07:00', '12:00')]);
  assert.equal(pDays[18].totalHours, 5);
  assert.equal(pDays[18].overtimeHours, 0);
  // 匯出：週六那列的班別起訖顯示 08:00／16:30，平日那列仍是 C 班 08:20／17:30
  load(c, [input(5, 'C', '07:48', '16:31'), input(7, 'C', '07:48', '16:31')]);
  callbacks.B();
  const sheet = exports[0].sheets[1];
  assert.deepEqual([sheet['6:2'].v, sheet['6:4'].v, sheet['6:5'].v], ['C', '08:00', '16:30']);   // 9/5
  assert.deepEqual([sheet['8:2'].v, sheet['8:4'].v, sheet['8:5'].v], ['C', '08:20', '17:30']);   // 9/7
});

test('day type: national holiday > Sunday > Saturday > weekday typhoon > workday', () => {
  const { context: c } = harness();
  const typhoon = { name: '颱風假', type: 'typhoon' }, national = { name: '國定假日', type: 'national' };
  const tue = new Date(2026, 8, 1), sat = new Date(2026, 8, 5), sun = new Date(2026, 8, 6);
  assert.equal(c.getDayType(tue, null), 'workday');
  assert.equal(c.getDayType(tue, typhoon), 'typhoon');
  assert.equal(c.getDayType(sat, typhoon), 'restDay');      // 颱風假遇週六 → 週六規則
  assert.equal(c.getDayType(sun, typhoon), 'restHoliday');  // 颱風假遇週日 → 週日規則
  assert.equal(c.getDayType(sat, national), 'restHoliday');
  assert.equal(c.getDayType(sun, null), 'restHoliday');
});

test('weekday typhoon day is not counted for attendance and earns no overtime', () => {
  const { context: c } = harness();
  const typhoonDays = [1, 2, 3, 5, 6];
  c.getHolidayForDate = (date) => (typhoonDays.includes(date.getDate()) ? { name: '颱風假', type: 'typhoon' } : null);
  const days = load(c, [
    input(1, 'A', '09:00', '19:30'),   // 平日颱風假有出勤：晚到、晚走、工時不足都不算
    input(2, 'A', '', ''),             // 平日颱風假沒出勤：不算未出勤
    input(3, 'A', '', '17:30'),        // 平日颱風假單邊打卡：不標異常
    input(5, 'A', '08:30', '17:30'),   // 颱風假遇週六 → 週六規則
    input(6, 'A', '08:30', '10:30')    // 颱風假遇週日 → 休假日規則
  ]);
  assert.equal(days[0].typhoonDay, true);
  assert.equal(days[0].overtimeHours, 0);
  assert.equal(days[0].overtimeMultiplied, 0);
  assert.equal(days[0].lateMinutes, 0);
  assert.equal(days[0].short, false);
  assert.equal(days[0].note, '颱風假出勤（不計考勤、不計加班）');
  assert.equal(c.getReportDayStatus(days[0], 54), '颱風假出勤');
  assert.equal(days[1].noShow, false);
  assert.equal(days[2].missingPunch, false);
  assert.equal(days[4].overtimeHours, 8);
  assert.equal(days[4].overtimeMultiplied, 12.7);
  assert.equal(days[5].overtimeHours, 8);
  assert.equal(days[5].overtimeMultiplied, 8);
  const a = c.getReport().people[0].anomalies;
  assert.deepEqual([a.lateCount, a.earlyCount, a.missingStart, a.missingEnd, a.shortDays, a.noShowDays], [0, 0, 0, 0, 0, 0]);
});

test('person shift table fills in shifts below the attendance file and manual choices', () => {
  const { context: c } = harness();
  c.personShifts = [{ no: '0002', name: '張慕真', shift: 'C' }, { no: '0004', name: '陳宗毅', shift: 'B' }];
  const mk = (day, extra) => ({ 姓名: ' 張慕真 ', 工號: '0002', 日期: new Date(2026, 8, day), 上班1: '08:20', 下班1: '17:30', ...extra });
  const days = load(c, [mk(1), mk(2, { 對應時段: 'B' }), mk(3, { 工號: '9999' })]);
  assert.equal(days[0].shift.name, 'C');      // 沒有對應時段 → 人員對應表
  assert.equal(days[1].shift.name, 'B');      // 考勤檔對應時段優先
  assert.equal(days[2].shift.name, 'C');      // 該員 9/3 沒有資料列，仍依人員對應表
  const other = c.getReport().people.find((p) => p.employeeNo === '9999');
  assert.equal(other.days[2].shift.name, 'A'); // 同名但編號不符 → 不採用，退回第一個班別
  c.shiftOverrides['0002|張慕真|2026/9/1'] = 'P';
  assert.equal(c.getReport().people[0].days[0].shift.name, 'P');
});

test('person shift sheet is parsed from the header row that has both name and shift', () => {
  const { context: c } = harness();
  const aoa = [
    ['', '班別', '工作時間', '', ''],
    ['', 'A', '08:30~12:30', '13:30~17:30', ''],
    ['', '', '', '', ''],
    ['員工\r\n編號', '姓名', '班別'],
    ['0002', ' 張慕真 ', 'A'],
    [2002, '陳春宏', 'C'],
    ['2029', '劉怡慧', ''],
    ['', '', ''],
    ['9999', '空白列之後不讀', 'A']
  ];
  c.XLSX.utils.sheet_to_json = () => aoa;
  const book = { SheetNames: ['s'], Sheets: { s: {} } };
  const list = Array.from(c.parsePersonShiftSheet(book), (x) => ({ ...x }));
  assert.deepEqual(list, [{ no: '0002', name: '張慕真', shift: 'A' }, { no: '2002', name: '陳春宏', shift: 'C' }]);
  c.XLSX.utils.sheet_to_json = () => [['姓名', '日期']];
  assert.deepEqual(Array.from(c.parsePersonShiftSheet(book)), []);
});

test('monthly overtime limit defaults to 54 and honors 0', () => {
  const { context: c } = harness();
  assert.equal(c.getOvertimeLimit(), 54);
  c.overtimeLimitInput.value = '';
  assert.equal(c.getOvertimeLimit(), 54);
  c.overtimeLimitInput.value = '0';
  assert.equal(c.getOvertimeLimit(), 0);
  c.overtimeLimitInput.value = '46';
  assert.equal(c.getOvertimeLimit(), 46);
  c.overtimeLimitInput.value = 'abc';
  assert.equal(c.getOvertimeLimit(), 54);
  assert.match(html, /id="overtimeLimit"[^>]*value="54"/);
});

// 瀏覽器版 SheetJS 沒有 Big5 編碼表，XLSX.read 的 codepage 選項會被默默忽略（Node 才會自動載入），
// 所以 .xls 必須在讀完後自行用 TextDecoder("big5") 還原；這裡用 2611.xls 實際讀出的亂碼字串驗證。
test('.xls strings are re-decoded from Big5 mojibake; other formats are left alone', () => {
  const { context: c } = harness();
  const mojibake = '¦Ò¶Ô¸¹½X';   // 「考勤號碼」被當成西歐編碼
  const make = () => ({
    SheetNames: ['s'],
    Sheets: { s: { '!ref': 'A1:C1', A1: { t: 's', v: mojibake, w: mojibake }, B1: { t: 's', v: 'Dale' }, C1: { t: 'n', v: 5 } } }
  });
  c.XLSX.read = () => make();
  const xls = c.readWorkbookFile(new ArrayBuffer(0), 'a.XLS');
  assert.equal(xls.Sheets.s.A1.v, '考勤號碼');
  assert.equal(xls.Sheets.s.A1.w, '考勤號碼');
  assert.equal(xls.Sheets.s.B1.v, 'Dale');
  assert.equal(xls.Sheets.s.C1.v, 5);
  ['b.xlsx', 'c.csv', ''].forEach((name) => assert.equal(c.readWorkbookFile(new ArrayBuffer(0), name).Sheets.s.A1.v, mojibake));
});

test('Big5 re-decoding skips text that is already correct Unicode', () => {
  const { context: c } = harness();
  assert.equal(c.redecodeBig5Text('考勤號碼'), '考勤號碼');
  assert.equal(c.redecodeBig5Text('Café 2026/11/1'), 'Café 2026/11/1');
  assert.equal(c.redecodeBig5Text(''), '');
  assert.equal(c.redecodeBig5Text(42), 42);
});

test('people sort by employee/attendance number numerically, then name; list shows the shift', () => {
  const { context: c } = harness();
  const mk = (no, name) => ({ 姓名: name, 工號: no, 日期: new Date(2026, 8, 1), 上班1: '08:30', 下班1: '17:30' });
  c.loadWorkbook({ SheetNames: ['d'], Sheets: { d: [mk('2002', '乙'), mk('0079', '丙'), mk('3009', '甲'), mk('2002', '丁'), mk('0002', '戊')] } });
  const people = c.getReport().people;
  assert.deepEqual(Array.from(people, (p) => p.employeeNo + p.name), ['0002戊', '0079丙', '2002乙', '2002丁', '3009甲']);
  assert.equal(c.personShiftLabel(people[0]), '班別 A');
  people[0].days[3] = { ...people[0].days[3], shift: { name: 'B' } };
  assert.equal(c.personShiftLabel(people[0]), '混合班別');
  assert.match(html, /<label class="month-shift-control hidden"/);
  assert.match(html, /id="complianceButton" class="button hidden"/);
  assert.match(html, /id="exportButton"[^>]*>考勤檢核報表</);
});

test('days without attendance rows use the person\'s usual shift, so the list is not "混合班別"', () => {
  const { context: c, callbacks, exports } = harness();
  // 測試乙只有 9/1～9/3 有資料（B 班），9/4 之後沒有資料、也沒有人員對應表。
  const days = load(c, [input(1, 'B'), input(2, 'B'), input(3, 'B')]);
  assert.equal(days[29].shift.name, 'B');                 // 9/30 無資料仍是 B，不退回第一個班別 A
  const person = c.getReport().people[0];
  assert.equal(c.personShiftLabel(person), '班別 B');
  callbacks.B();
  const sheet = exports[0].sheets[1];
  assert.equal(sheet['31:2'].v, 'B');                     // 匯出 9/30 那列的班別
  assert.equal(sheet['31:4'].v, '08:30');
  assert.equal(sheet['31:5'].v, '17:00');                 // B 班的下班時間，不是 A 班的 17:30
});

test('usual shift: most frequent recognised period wins; the person table still takes priority', () => {
  const { context: c } = harness();
  const person = { rows: [{ shiftName: 'C' }, { shiftName: 'B' }, { shiftName: 'C' }, { shiftName: '未知' }, { shiftName: '' }] };
  assert.equal(c.getPersonUsualShiftName(person), 'C');
  assert.equal(c.getPersonUsualShiftName({ rows: [{ shiftName: 'B' }, { shiftName: 'C' }] }), 'B');  // 同次數取先出現
  assert.equal(c.getPersonUsualShiftName({ rows: [] }), '');
  c.personShifts = [{ no: 'T1', name: '測試員工', shift: 'C' }];
  const days = load(c, [input(1, 'B')]);
  assert.equal(days[0].shift.name, 'B');                  // 當天對應時段最優先
  assert.equal(days[10].shift.name, 'C');                 // 無資料日：人員對應表優先於當月推斷
});

test('late / early are computed from the shift on workdays only, with 0 grace minutes', () => {
  const { context: c } = harness();
  const [a, b, cShift, p] = c.getAvailableShifts();
  assert.deepEqual({ ...c.getLateEarlyMinutes({ start1: '08:31', end1: '17:30' }, a) }, { lateMinutes: 1, earlyMinutes: 0 });
  assert.deepEqual({ ...c.getLateEarlyMinutes({ start1: '08:30', end1: '16:59' }, b) }, { lateMinutes: 0, earlyMinutes: 1 });
  assert.deepEqual({ ...c.getLateEarlyMinutes({ start1: '08:21', end1: '12:26' }, cShift) }, { lateMinutes: 1, earlyMinutes: 304 });
  assert.deepEqual({ ...c.getLateEarlyMinutes({ start1: '', end1: '17:00' }, a) }, { lateMinutes: 0, earlyMinutes: 30 });   // 只有下班卡：只算早退
  assert.deepEqual({ ...c.getLateEarlyMinutes({ start1: '09:00', end1: '' }, a) }, { lateMinutes: 30, earlyMinutes: 0 });   // 只有上班卡：只算遲到
  assert.deepEqual({ ...c.getLateEarlyMinutes({ start1: '10:00', end1: '12:00' }, p) }, { lateMinutes: 0, earlyMinutes: 0 }); // P 班不算
  // 9/5 週六晚到、9/2 國定假日晚到都不算遲到
  c.getHolidayForDate = (date) => (date.getDate() === 2 ? { name: '測試假日', type: 'national' } : null);
  const days = load(c, [input(1, 'A', '08:45', '17:10'), input(2, 'A', '10:00', '17:30'), input(5, 'A', '10:00', '17:30')]);
  assert.equal(days[0].lateMinutes, 15);
  assert.equal(days[0].earlyMinutes, 20);
  assert.equal(days[0].note, '未滿 8 小時；遲到 15 分；早退 20 分');
  assert.equal(days[1].lateMinutes, 0);
  assert.equal(days[4].lateMinutes, 0);
});

test('no-show counts only workdays that have a row with no punches at all', () => {
  const { context: c } = harness();
  // 9/1 週二整天沒打卡（有資料列）→ 未出勤；9/5 週六沒打卡 → 不算；9/2～9/4、9/7 之後沒有資料列 → 不算。
  const days = load(c, [input(1, 'A', '', ''), input(5, 'A', '', ''), input(8, 'A', '08:30', '08:30')]);
  assert.equal(days[0].noShow, true);
  assert.equal(days[0].note, '未出勤（請假／曠職待確認）');
  assert.equal(days[4].noShow, false);
  assert.equal(days[2].noShow, false);
  assert.equal(days[7].noShow, false);   // 有打卡但工時 0，不是未出勤
  assert.equal(days[7].earlyMinutes, 540); // 但仍算早退（17:30 − 08:30），不能從異常匯總消失
  assert.equal(days[7].note, '工時 0；早退 540 分');
});

test('anomaly summary appears under each person on screen data and in both export places', () => {
  const { context: c, callbacks, exports } = harness();
  load(c, [
    input(1, 'A', '08:45', '17:30'),   // 遲到 15
    input(2, 'A', '08:40', '17:00'),   // 遲到 10、早退 30、未滿
    input(3, 'A', '', '17:30'),        // 上班未打卡
    input(4, 'A', '08:30', ''),        // 下班未打卡
    input(7, 'A', '', '')              // 未出勤
  ]);
  const person = c.getReport().people[0];
  assert.deepEqual({ ...person.anomalies }, { lateCount: 2, lateMinutes: 25, earlyCount: 1, earlyMinutes: 30, missingStart: 1, missingEnd: 1, shortDays: 2, noShowDays: 1 });
  assert.deepEqual(Array.from(c.anomalyItems(person.anomalies), (x) => Array.from(x)), [
    ['遲到', '2 次，共 25 分鐘'], ['早退', '1 次，共 30 分鐘'], ['上班未打卡', '1 次'], ['下班未打卡', '1 次'],
    ['未滿時數', '2 天'], ['未出勤（請假／曠職待確認）', '1 天']
  ]);
  callbacks.B();
  const [overview, sheet] = exports[0].sheets;
  assert.equal(overview['1:4'].v, '遲到次數');
  assert.deepEqual([2, 3, 4, 5, 6, 7, 8, 9, 10, 11].map((col) => overview[`2:${col}`].v), ['測試員工', 'A', 2, 25, 1, 30, 1, 1, 2, 1]);
  // 個人表：30 天（第 2～31 列）→ 兩列加班合計（32、33）→ 異常匯總標題（34）→ 6 項（35～40）
  assert.equal(sheet['34:0'].v, '異常匯總');
  assert.equal(sheet['35:0'].v, '遲到');
  assert.equal(sheet['35:2'].v, '2 次，共 25 分鐘');
  assert.equal(sheet['40:0'].v, '未出勤（請假／曠職待確認）');
  assert.equal(sheet['40:2'].v, '1 天');
});

test('export file is named after the button and the selected month', () => {
  const { context: c, callbacks } = harness();
  const links = [];
  c.document.createElement = () => { const link = { click() {} }; links.push(link); return link; };
  load(c, [input(1, 'A')]);
  callbacks.B();
  assert.equal(links[0].download, '考勤檢核報表_2026-09.xlsx');
});

// 只能釘住設定存在；實際不溢出要在瀏覽器各寬度量測（Node 沒有版面引擎）。
test('page grid column cannot be stretched by the 1080px detail table', () => {
  assert.match(html, /\.app \{[^}]*grid-template-columns: minmax\(0, 1fr\)/);
  assert.match(html, /main \{[^}]*min-width: 0/);
});

test('embedded JavaScript parses', () => {
  new vm.Script(html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')));
});
