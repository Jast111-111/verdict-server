// ============================================================
//  ВЕРДИКТ — сервер экспертной оценки
//  Держит настройки, список участников и оценки в памяти,
//  сохраняет на диск в data/store.json (переживает перезапуск)
//  и в data/Оценки.xlsx (стандартный Excel-файл, обновляется
//  автоматически после каждой новой оценки).
// ============================================================

const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const os = require('os');
const multer = require('multer');
const mammoth = require('mammoth');

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const STORE_FILE = path.join(DATA_DIR, 'store.json');
const XLSX_FILE = path.join(DATA_DIR, 'Оценки.xlsx');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const DEFAULT_CRITERIA = ['Критерий 1', 'Критерий 2', 'Критерий 3', 'Критерий 4'];
const DEFAULT_EXPERTS = ['Эксперт 1', 'Эксперт 2', 'Эксперт 3', 'Эксперт 4'];
const MODERATOR_PASSWORD = process.env.MODERATOR_PASSWORD || 'вердикт2026';

let store = {
  config: { eventName: 'Мероприятие', criteria: [...DEFAULT_CRITERIA], experts: [...DEFAULT_EXPERTS] },
  participants: [],
  submissions: [],
};

if (fs.existsSync(STORE_FILE)) {
  try {
    const loaded = JSON.parse(fs.readFileSync(STORE_FILE, 'utf8'));
    store = { ...store, ...loaded };
  } catch (e) {
    console.error('Не удалось прочитать data/store.json, начинаю с чистого состояния:', e.message);
  }
}
if (!Array.isArray(store.config.experts) || store.config.experts.length < 1) store.config.experts = [...DEFAULT_EXPERTS];
if (!Array.isArray(store.participants)) store.participants = [];
if (!Array.isArray(store.submissions)) store.submissions = [];

function persist() {
  fs.writeFileSync(STORE_FILE, JSON.stringify(store, null, 2), 'utf8');
}

function nextParticipantNumber() {
  const nums = store.participants.map((p) => parseInt(p.number, 10)).filter((n) => !isNaN(n));
  return (nums.length ? Math.max(...nums) : 0) + 1;
}

// ---------- разбор загруженного файла со списком участников ----------
// Ожидаемые столбцы (без заголовка или с ним — заголовок определяется
// автоматически): Имя | Возраст | Мероприятие/номинация
function extractTableRowsFromHtml(html) {
  const rows = [];
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/g;
  let trMatch;
  while ((trMatch = trRe.exec(html))) {
    const cells = [];
    const tdRe = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g;
    let tdMatch;
    while ((tdMatch = tdRe.exec(trMatch[1]))) {
      const text = tdMatch[1]
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/\s+/g, ' ')
        .trim();
      cells.push(text);
    }
    if (cells.length) rows.push(cells);
  }
  return rows;
}

async function parseParticipantsFile(buffer, originalName) {
  const ext = (originalName.split('.').pop() || '').toLowerCase();
  let rows = [];

  if (ext === 'xlsx' || ext === 'xls') {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const sheet = wb.worksheets[0];
    if (sheet) {
      sheet.eachRow((row) => {
        const vals = row.values.slice(1).map((v) => (v === null || v === undefined ? '' : String(v).trim()));
        if (vals.some((v) => v)) rows.push(vals);
      });
    }
  } else if (ext === 'docx') {
    const result = await mammoth.convertToHtml({ buffer });
    rows = extractTableRowsFromHtml(result.value);
    if (rows.length === 0) {
      // нет таблицы — пробуем построчный текст "Имя, Возраст, Мероприятие" или "Имя — Возраст — Мероприятие"
      const text = result.value.replace(/<[^>]+>/g, '\n');
      rows = text
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean)
        .map((l) => l.split(/[,;|—-]\s*/).map((p) => p.trim()));
    }
  } else {
    throw new Error('Поддерживаются только файлы .xlsx и .docx');
  }

  if (rows.length === 0) return [];

  // если первая строка похожа на заголовок (возраст не число) — пропускаем её
  const first = rows[0];
  const looksLikeHeader = first[1] !== undefined && first[1] !== '' && isNaN(parseInt(first[1], 10));
  const dataRows = looksLikeHeader ? rows.slice(1) : rows;

  let n = nextParticipantNumber();
  return dataRows
    .filter((r) => r[0] && r[0].trim())
    .map((r) => {
      const ageNum = r[1] !== undefined ? parseInt(r[1], 10) : NaN;
      return {
        id: Date.now() + '_' + Math.random().toString(36).slice(2, 8),
        number: String(n++),
        name: r[0].trim().slice(0, 150),
        age: isNaN(ageNum) ? null : ageNum,
        event: (r[2] || '').trim().slice(0, 200),
      };
    });
}

// ---------- генерация Excel ----------
const THIN = { style: 'thin', color: { argb: 'FF000000' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

async function writeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Отчёт по участникам');

  const crit = store.config.criteria;
  const nCrit = crit.length;
  const lastCol = 4 + nCrit;

  ws.getColumn(1).width = 26;
  for (let i = 0; i < nCrit; i++) ws.getColumn(2 + i).width = 12;
  ws.getColumn(2 + nCrit).width = 10;
  ws.getColumn(3 + nCrit).width = 12;
  ws.getColumn(4 + nCrit).width = 12;

  const groups = [];
  const byKey = new Map();
  store.submissions.forEach((s) => {
    const key = s.participantId || s.target;
    if (!byKey.has(key)) { byKey.set(key, groups.length); groups.push({ ...s, items: [] }); }
    groups[byKey.get(key)].items.push(s);
  });

  let r = 1;
  if (groups.length === 0) {
    ws.getCell(1, 1).value = 'Пока нет ни одной оценки';
    ws.getCell(1, 1).font = { italic: true };
  }

  groups.forEach((g) => {
    const rows = g.items;
    const infoRows = [
      ['УЧАСТНИК №:', g.participantNumber || '—'],
      ['ФИО / НАЗВАНИЕ:', g.target],
      ['ВОЗРАСТ:', g.participantAge != null ? String(g.participantAge) : '—'],
      ['МЕРОПРИЯТИЕ / НОМИНАЦИЯ:', g.targetGroup || '—'],
    ];
    infoRows.forEach(([label, value]) => {
      ws.mergeCells(r, 1, r, 2);
      ws.getCell(r, 1).value = label;
      ws.getCell(r, 1).font = { bold: true };
      for (let c = 1; c <= 2; c++) ws.getCell(r, c).border = BORDER;
      ws.mergeCells(r, 3, r, lastCol);
      ws.getCell(r, 3).value = value;
      ws.getCell(r, 3).font = { bold: true };
      for (let c = 3; c <= lastCol; c++) ws.getCell(r, c).border = BORDER;
      r++;
    });

    ws.getCell(r, 1).value = 'ЧЛЕН ЖЮРИ';
    ws.getCell(r, 1).font = { bold: true };
    ws.getCell(r, 1).border = BORDER;
    ws.mergeCells(r, 2, r, 1 + nCrit);
    ws.getCell(r, 2).value = 'Оценки';
    ws.getCell(r, 2).font = { bold: true };
    ws.getCell(r, 2).alignment = { horizontal: 'center' };
    for (let c = 2; c <= 1 + nCrit; c++) ws.getCell(r, c).border = BORDER;
    ws.getCell(r, 2 + nCrit).value = 'Итого';
    ws.getCell(r, 2 + nCrit).font = { bold: true };
    ws.getCell(r, 2 + nCrit).border = BORDER;
    r++;

    const firstJudgeRow = r;
    rows.forEach((s) => {
      ws.getCell(r, 1).value = s.expertName;
      ws.getCell(r, 1).border = BORDER;
      s.scores.forEach((sc, ci) => {
        const cell = ws.getCell(r, 2 + ci);
        cell.value = sc.score;
        cell.alignment = { horizontal: 'center' };
        cell.font = { bold: true };
        cell.border = BORDER;
      });
      const totalCell = ws.getCell(r, 2 + nCrit);
      totalCell.value = s.total;
      totalCell.alignment = { horizontal: 'center' };
      totalCell.font = { bold: true };
      totalCell.border = BORDER;
      r++;
    });
    const lastJudgeRow = r - 1;

    const grandTotal = rows.reduce((a, s) => a + s.total, 0);
    ws.mergeCells(firstJudgeRow, 3 + nCrit, lastJudgeRow, 3 + nCrit);
    const sumLbl = ws.getCell(firstJudgeRow, 3 + nCrit);
    sumLbl.value = 'СУММА:';
    sumLbl.font = { bold: true };
    sumLbl.alignment = { vertical: 'middle', horizontal: 'center' };
    sumLbl.border = BORDER;
    ws.mergeCells(firstJudgeRow, 4 + nCrit, lastJudgeRow, 4 + nCrit);
    const sumVal = ws.getCell(firstJudgeRow, 4 + nCrit);
    sumVal.value = grandTotal;
    sumVal.font = { bold: true, size: 16 };
    sumVal.alignment = { vertical: 'middle', horizontal: 'center' };
    const medium = { style: 'medium', color: { argb: 'FF000000' } };
    sumVal.border = { top: medium, left: medium, bottom: medium, right: medium };

    r++;
    ws.mergeCells(r, 2, r, lastCol);
    ws.getCell(r, 2).value = 'КОММЕНТАРИИ:';
    ws.getCell(r, 2).font = { bold: true };
    r++;

    rows.forEach((s) => {
      ws.getCell(r, 1).value = s.expertName;
      ws.getCell(r, 1).font = { bold: true };
      ws.getCell(r, 1).border = BORDER;
      ws.getCell(r, 1).alignment = { wrapText: true, vertical: 'top' };
      ws.mergeCells(r, 2, r, lastCol);
      const cCell = ws.getCell(r, 2);
      cCell.value = s.comment || '';
      cCell.alignment = { wrapText: true, vertical: 'top' };
      cCell.border = BORDER;
      ws.getRow(r).height = 32;
      r++;
    });

    r += 2;
  });

  // сводный лист со всеми оценками построчно
  const ws2 = wb.addWorksheet('Все оценки списком');
  const cols2 = [
    { header: '#', key: 'n', width: 5 },
    { header: 'Время', key: 'time', width: 18 },
    { header: 'Эксперт', key: 'expert', width: 20 },
    { header: '№ участника', key: 'num', width: 12 },
    { header: 'ФИО / название', key: 'target', width: 22 },
    { header: 'Возраст', key: 'age', width: 10 },
    { header: 'Мероприятие', key: 'group', width: 20 },
  ];
  crit.forEach((c, i) => cols2.push({ header: c, key: 'score' + i, width: 12 }));
  cols2.push({ header: 'Итого', key: 'total', width: 10 });
  cols2.push({ header: 'Комментарий', key: 'comment', width: 40 });
  ws2.columns = cols2;
  ws2.getRow(1).font = { bold: true };
  ws2.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFE7CE' } };
  ws2.views = [{ state: 'frozen', ySplit: 1 }];
  store.submissions.forEach((s, i) => {
    const row = {
      n: i + 1,
      time: new Date(s.timestamp).toLocaleString('ru-RU'),
      expert: s.expertName,
      num: s.participantNumber || '',
      target: s.target,
      age: s.participantAge != null ? s.participantAge : '',
      group: s.targetGroup || '',
      total: s.total,
      comment: s.comment || '',
    };
    s.scores.forEach((sc, idx) => { row['score' + idx] = sc.score; });
    ws2.addRow(row);
  });

  await wb.xlsx.writeFile(XLSX_FILE);
}

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8 * 1024 * 1024 } });

app.get('/export.xlsx', async (req, res) => {
  try {
    await writeXlsx();
    res.download(XLSX_FILE, `Оценки_${new Date().toISOString().slice(0, 10)}.xlsx`);
  } catch (e) {
    console.error(e);
    res.status(500).send('Не удалось сформировать файл');
  }
});

const server = http.createServer(app);
const io = new Server(server);

app.post('/api/import-participants', upload.single('file'), async (req, res) => {
  try {
    if (String(req.body.password || '') !== MODERATOR_PASSWORD) {
      return res.status(401).json({ ok: false, error: 'Неверный пароль' });
    }
    if (!req.file) return res.status(400).json({ ok: false, error: 'Файл не получен' });
    const added = await parseParticipantsFile(req.file.buffer, req.file.originalname);
    if (added.length === 0) {
      return res.json({ ok: true, added: 0, message: 'В файле не найдено ни одной строки с именем участника' });
    }
    store.participants.push(...added);
    persist();
    io.emit('participants:update', store.participants);
    res.json({ ok: true, added: added.length });
  } catch (e) {
    console.error('Ошибка импорта участников:', e);
    res.status(500).json({ ok: false, error: e.message || 'Ошибка обработки файла' });
  }
});

io.on('connection', (socket) => {
  socket.emit('init', store);

  socket.on('moderator:login', (password, cb) => {
    const ok = String(password || '') === MODERATOR_PASSWORD;
    if (typeof cb === 'function') cb({ ok });
  });

  socket.on('config:save', (cfg, cb) => {
    if (!cfg || String(cfg.password || '') !== MODERATOR_PASSWORD) {
      if (typeof cb === 'function') cb({ ok: false });
      return;
    }
    store.config = {
      eventName: String(cfg.eventName || 'Мероприятие').slice(0, 200),
      criteria: Array.isArray(cfg.criteria) && cfg.criteria.length === 4
        ? cfg.criteria.map((c) => String(c || '').slice(0, 100))
        : store.config.criteria,
      experts: Array.isArray(cfg.experts) && cfg.experts.length >= 1
        ? cfg.experts.map((e) => String(e || '').slice(0, 100)).slice(0, 60)
        : store.config.experts,
    };
    persist();
    io.emit('config:update', store.config);
    if (typeof cb === 'function') cb({ ok: true });
  });

  // ---------- участники: добавление/редактирование/удаление вручную ----------
  socket.on('participants:add', (payload, cb) => {
    if (!payload || String(payload.password || '') !== MODERATOR_PASSWORD) {
      if (typeof cb === 'function') cb({ ok: false });
      return;
    }
    const p = payload.participant || {};
    if (!p.name || !String(p.name).trim()) {
      if (typeof cb === 'function') cb({ ok: false, error: 'Укажите имя участника' });
      return;
    }
    const ageNum = parseInt(p.age, 10);
    store.participants.push({
      id: Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      number: String(p.number || nextParticipantNumber()),
      name: String(p.name).trim().slice(0, 150),
      age: isNaN(ageNum) ? null : ageNum,
      event: String(p.event || '').trim().slice(0, 200),
    });
    persist();
    io.emit('participants:update', store.participants);
    if (typeof cb === 'function') cb({ ok: true });
  });

  socket.on('participants:edit', (payload, cb) => {
    if (!payload || String(payload.password || '') !== MODERATOR_PASSWORD) {
      if (typeof cb === 'function') cb({ ok: false });
      return;
    }
    const idx = store.participants.findIndex((p) => p.id === payload.id);
    if (idx === -1) { if (typeof cb === 'function') cb({ ok: false }); return; }
    const patch = payload.patch || {};
    const ageNum = parseInt(patch.age, 10);
    store.participants[idx] = {
      ...store.participants[idx],
      number: patch.number !== undefined ? String(patch.number) : store.participants[idx].number,
      name: patch.name !== undefined ? String(patch.name).trim().slice(0, 150) : store.participants[idx].name,
      age: patch.age !== undefined ? (isNaN(ageNum) ? null : ageNum) : store.participants[idx].age,
      event: patch.event !== undefined ? String(patch.event).trim().slice(0, 200) : store.participants[idx].event,
    };
    persist();
    io.emit('participants:update', store.participants);
    if (typeof cb === 'function') cb({ ok: true });
  });

  socket.on('participants:remove', (payload, cb) => {
    if (!payload || String(payload.password || '') !== MODERATOR_PASSWORD) {
      if (typeof cb === 'function') cb({ ok: false });
      return;
    }
    store.participants = store.participants.filter((p) => p.id !== payload.id);
    persist();
    io.emit('participants:update', store.participants);
    if (typeof cb === 'function') cb({ ok: true });
  });

  // ---------- оценки ----------
  socket.on('submission:add', (sub) => {
    if (!sub || !sub.expertName || !Array.isArray(sub.scores)) return;
    if (sub.scores.length !== store.config.criteria.length) return;
    if (String(sub.comment || '').trim().length === 0) return;
    const participant = store.participants.find((p) => p.id === sub.participantId);
    if (!participant) return;

    const clean = {
      id: Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      expertId: Number.isInteger(sub.expertId) ? sub.expertId : null,
      expertName: String(sub.expertName).slice(0, 100),
      participantId: participant.id,
      participantNumber: participant.number,
      target: participant.name,
      targetGroup: participant.event || '',
      participantAge: participant.age != null ? participant.age : null,
      scores: sub.scores.map((s) => ({
        criterion: String(s.criterion || '').slice(0, 100),
        score: Math.max(1, Math.min(10, Math.round(Number(s.score) || 1))),
      })),
      comment: String(sub.comment || '').slice(0, 2000),
      timestamp: new Date().toISOString(),
    };
    clean.total = clean.scores.reduce((a, c) => a + c.score, 0);
    clean.average = +(clean.total / clean.scores.length).toFixed(1);

    store.submissions.push(clean);
    persist();
    writeXlsx().catch((e) => console.error('Ошибка автосохранения xlsx:', e.message));
    io.emit('submissions:update', store.submissions);
  });

  socket.on('data:clear', (password, cb) => {
    if (String(password || '') !== MODERATOR_PASSWORD) {
      if (typeof cb === 'function') cb({ ok: false });
      return;
    }
    store.submissions = [];
    persist();
    writeXlsx().catch(() => {});
    io.emit('submissions:update', store.submissions);
    if (typeof cb === 'function') cb({ ok: true });
  });
});

function localIPs() {
  const nets = os.networkInterfaces();
  const results = [];
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === 'IPv4' && !net.internal) results.push(net.address);
    }
  }
  return results;
}

server.listen(PORT, () => {
  console.log('==============================================================');
  console.log('  ВЕРДИКТ — сервер запущен');
  console.log('  Открыть на этом компьютере:      http://localhost:' + PORT);
  localIPs().forEach((ip) => console.log('  В локальной сети доступен по:     http://' + ip + ':' + PORT));
  console.log('');
  console.log('  Чтобы дать доступ через интернет (без общей сети), в НОВОМ');
  console.log('  окне терминала, не закрывая это, выполните:');
  console.log('      npm run tunnel');
  console.log('  и разошлите экспертам ссылку, которую он выведет.');
  console.log('');
  console.log('  Пароль входа для модератора: ' + MODERATOR_PASSWORD);
  console.log('  (задать свой пароль: переменная окружения MODERATOR_PASSWORD)');
  console.log('==============================================================');
});
