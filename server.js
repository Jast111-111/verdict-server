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

// ---------- генерация Excel: один номер участника = один лист ----------
const THIN = { style: 'thin', color: { argb: 'FF000000' } };
const MEDIUM = { style: 'medium', color: { argb: 'FF000000' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const BORDER_MED = { top: MEDIUM, left: MEDIUM, bottom: MEDIUM, right: MEDIUM };

function sanitizeSheetName(name, used) {
  let base = String(name || 'Лист').replace(/[\\/*?:[\]]/g, ' ').trim().slice(0, 28) || 'Лист';
  let final = base, n = 2;
  while (used.has(final)) { final = (base.slice(0, 25) + ' ' + n).trim(); n++; }
  used.add(final);
  return final;
}

// строит один лист-ведомость на участника (УЧАСТНИК / КОНКУРСНЫЙ НОМЕР / ЧЛЕН ЖЮРИ×Оценки / СУММА / КОММЕНТАРИИ)
function addParticipantSheet(wb, p, usedNames, firstCritCol, lastCritCol, sumCol, commentEndCol, experts) {
  const nExperts = experts.length;
  const subsByExpert = new Map();
  store.submissions
    .filter((s) => s.participantId === p.id)
    .forEach((s) => subsByExpert.set(s.expertId, s));

  const sheetName = sanitizeSheetName(p.number, usedNames);
  const ws = wb.addWorksheet(sheetName);

  ws.getColumn(2).width = 22; ws.getColumn(3).width = 14; ws.getColumn(4).width = 14;
  for (let c = firstCritCol; c <= lastCritCol; c++) ws.getColumn(c).width = 10;
  ws.getColumn(sumCol - 1).width = 4;
  ws.getColumn(sumCol).width = 14;

  const setBox = (r1, c1, r2, c2, value, opts = {}) => {
    if (!(r1 === r2 && c1 === c2)) ws.mergeCells(r1, c1, r2, c2);
    const cell = ws.getCell(r1, c1);
    cell.value = value;
    if (opts.bold) cell.font = { bold: true, size: opts.size };
    if (opts.align) cell.alignment = { horizontal: opts.align, vertical: 'middle', wrapText: !!opts.wrap };
    for (let rr = r1; rr <= r2; rr++) for (let cc = c1; cc <= c2; cc++) ws.getCell(rr, cc).border = opts.border || BORDER;
    return cell;
  };

  setBox(7, 2, 8, 4, 'УЧАСТНИК:', { bold: true, align: 'right' });
  setBox(7, 6, 8, 10, p.name, { bold: true });
  setBox(9, 2, 10, 4, 'КОНКУРСНЫЙ НОМЕР:', { bold: true, align: 'right' });
  setBox(9, 6, 10, 10, p.event || '—', { bold: true });

  setBox(12, 2, 12, 4, 'ЧЛЕН ЖЮРИ', { bold: true });
  setBox(12, firstCritCol, 12, lastCritCol, 'Оценки', { bold: true, align: 'center' });

  const firstJudgeRow = 13;
  experts.forEach((expertName, i) => {
    const r = firstJudgeRow + i;
    setBox(r, 2, r, 4, expertName, {});
    const sub = subsByExpert.get(i);
    if (sub) {
      sub.scores.forEach((sc, ci) => {
        setBox(r, firstCritCol + ci, r, firstCritCol + ci, sc.score, { bold: true, align: 'center' });
      });
    } else {
      for (let c = firstCritCol; c <= lastCritCol; c++) ws.getCell(r, c).border = BORDER;
    }
  });
  const lastJudgeRow = firstJudgeRow + nExperts - 1;

  const sumLblRow = nExperts >= 2 ? lastJudgeRow - 1 : firstJudgeRow;
  const sumValRow = lastJudgeRow;
  ws.getCell(sumLblRow, sumCol).value = 'СУММА:';
  ws.getCell(sumLblRow, sumCol).font = { bold: true };
  const totalCell = ws.getCell(sumValRow, sumCol);
  totalCell.value = { formula: `SUM(${ws.getCell(firstJudgeRow, firstCritCol).address}:${ws.getCell(lastJudgeRow, lastCritCol).address})` };
  totalCell.font = { bold: true, size: 16 };
  totalCell.alignment = { horizontal: 'center' };
  totalCell.border = BORDER_MED;

  let r = lastJudgeRow + 2;
  setBox(r, 3, r, lastCritCol, '', { border: BORDER });
  r++;
  setBox(r, 3, r, lastCritCol, 'КОММЕНТАРИИ:', { bold: true });
  r++;
  experts.forEach((expertName, i) => {
    const sub = subsByExpert.get(i);
    setBox(r, 2, r, 3, expertName, { wrap: true });
    setBox(r, 4, r, commentEndCol, sub ? (sub.comment || '') : '', { wrap: true });
    ws.getRow(r).height = 34;
    r++;
  });

  const total = [...subsByExpert.values()].reduce((a, s) => a + s.total, 0);
  return { number: p.number, name: p.name, event: p.event || '', age: p.age, total, hasAny: subsByExpert.size > 0 };
}

// собирает книгу: лист на каждого участника + опционально сводный лист "ИТОГО"
// порядок столбцов на сводном листе: Номер, Имя, Конкурсный номер, Возраст, Сумма баллов
async function buildParticipantsWorkbook(participantList, { withSummary = true } = {}) {
  const wb = new ExcelJS.Workbook();
  const crit = store.config.criteria;
  const nCrit = crit.length;
  const experts = store.config.experts;

  const firstCritCol = 5; // E
  const lastCritCol = firstCritCol + nCrit - 1;
  const sumCol = lastCritCol + 2;
  const commentEndCol = sumCol;

  const usedNames = new Set();
  const summaryRows = participantList.map((p) =>
    addParticipantSheet(wb, p, usedNames, firstCritCol, lastCritCol, sumCol, commentEndCol, experts)
  );

  if (withSummary) {
    const wsT = wb.addWorksheet(sanitizeSheetName('ИТОГО', usedNames));
    wsT.columns = [
      { header: 'Номер', key: 'number', width: 10 },
      { header: 'Имя', key: 'name', width: 32 },
      { header: 'Конкурсный номер', key: 'event', width: 26 },
      { header: 'Возраст', key: 'age', width: 10 },
      { header: 'Сумма баллов', key: 'total', width: 14 },
    ];
    wsT.getRow(1).font = { bold: true };
    wsT.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFE7CE' } };
    for (let c = 1; c <= 5; c++) wsT.getCell(1, c).border = BORDER;
    wsT.views = [{ state: 'frozen', ySplit: 1 }];
    summaryRows.forEach((row) => {
      const added = wsT.addRow({
        number: row.number,
        name: row.name,
        event: row.event,
        age: row.age != null ? row.age : '',
        total: row.hasAny ? row.total : '',
      });
      added.eachCell((cell) => { cell.border = BORDER; });
      added.getCell(5).alignment = { horizontal: 'center' };
      added.getCell(5).font = { bold: true };
    });
  }

  return wb;
}

// пересобирает основной файл на диске (автосохранение после каждой оценки)
async function writeXlsx() {
  const wb = await buildParticipantsWorkbook(store.participants, { withSummary: true });
  await wb.xlsx.writeFile(XLSX_FILE);
}

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8 * 1024 * 1024 } });

app.get('/export.xlsx', async (req, res) => {
  try {
    const participantId = req.query.participantId;
    if (participantId) {
      const p = store.participants.find((x) => x.id === participantId);
      if (!p) return res.status(404).send('Участник не найден');
      const wb = await buildParticipantsWorkbook([p], { withSummary: false });
      const fname = `Оценки_${p.name.replace(/[^\wа-яА-ЯёЁ]+/g, '_')}.xlsx`;
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fname)}"`);
      await wb.xlsx.write(res);
      return res.end();
    }
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
  socket.on('submission:add', (sub, cb) => {
    const fail = (error) => { if (typeof cb === 'function') cb({ ok: false, error }); };
    if (!sub || !sub.expertName || !Array.isArray(sub.scores)) return fail('Некорректные данные');
    if (sub.scores.length !== store.config.criteria.length) return fail('Некорректные данные');
    if (String(sub.comment || '').trim().length === 0) return fail('Комментарий обязателен');
    const participant = store.participants.find((p) => p.id === sub.participantId);
    if (!participant) return fail('Участник не найден');

    const already = store.submissions.some((s) => s.expertId === sub.expertId && s.participantId === participant.id);
    if (already) return fail('Вы уже оценили этого участника — оценка заблокирована для изменения');

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
    if (typeof cb === 'function') cb({ ok: true });
  });

  // модератор может исправить уже отправленную оценку/комментарий (опечатка и т.п.)
  socket.on('submission:edit', (payload, cb) => {
    const fail = (error) => { if (typeof cb === 'function') cb({ ok: false, error }); };
    if (!payload || String(payload.password || '') !== MODERATOR_PASSWORD) return fail('Неверный пароль модератора');
    const idx = store.submissions.findIndex((s) => s.id === payload.id);
    if (idx === -1) return fail('Оценка не найдена');
    const scoresIn = Array.isArray(payload.scores) ? payload.scores : null;
    if (!scoresIn || scoresIn.length !== store.config.criteria.length) return fail('Некорректные баллы');
    if (String(payload.comment || '').trim().length === 0) return fail('Комментарий обязателен');

    const s = store.submissions[idx];
    s.scores = scoresIn.map((sc, i) => ({
      criterion: s.scores[i] ? s.scores[i].criterion : store.config.criteria[i],
      score: Math.max(1, Math.min(10, Math.round(Number(sc.score ?? sc) || 1))),
    }));
    s.comment = String(payload.comment).slice(0, 2000);
    s.total = s.scores.reduce((a, c) => a + c.score, 0);
    s.average = +(s.total / s.scores.length).toFixed(1);
    s.editedAt = new Date().toISOString();

    persist();
    writeXlsx().catch((e) => console.error('Ошибка автосохранения xlsx:', e.message));
    io.emit('submissions:update', store.submissions);
    if (typeof cb === 'function') cb({ ok: true });
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
