// ============================================================
//  ВЕРДИКТ — сервер экспертной оценки
//  Держит текущее состояние (настройки + оценки) в памяти,
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

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const STORE_FILE = path.join(DATA_DIR, 'store.json');
const XLSX_FILE = path.join(DATA_DIR, 'Оценки.xlsx');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const DEFAULT_CRITERIA = ['Критерий 1', 'Критерий 2', 'Критерий 3', 'Критерий 4'];
const DEFAULT_EXPERTS = ['Эксперт 1', 'Эксперт 2', 'Эксперт 3', 'Эксперт 4'];
const MODERATOR_PASSWORD = process.env.MODERATOR_PASSWORD || 'вердикт2026';

let store = {
  config: { eventName: 'Мероприятие', criteria: [...DEFAULT_CRITERIA], experts: [...DEFAULT_EXPERTS], target: '', targetGroup: '' },
  submissions: [],
};

// на случай, если store.json был сохранён старой версией сервера — дополняем недостающее
if (!Array.isArray(store.config.experts) || store.config.experts.length < 1) {
  store.config.experts = [...DEFAULT_EXPERTS];
}
if (typeof store.config.target !== 'string') {
  store.config.target = '';
}
if (typeof store.config.targetGroup !== 'string') {
  store.config.targetGroup = '';
}

if (fs.existsSync(STORE_FILE)) {
  try {
    store = JSON.parse(fs.readFileSync(STORE_FILE, 'utf8'));
  } catch (e) {
    console.error('Не удалось прочитать data/store.json, начинаю с чистого состояния:', e.message);
  }
}

function persist() {
  fs.writeFileSync(STORE_FILE, JSON.stringify(store, null, 2), 'utf8');
}

async function writeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Отчёт по участникам');

  const crit = store.config.criteria;
  const nCrit = crit.length;
  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const lastCol = 4 + nCrit; // ЧЛЕН ЖЮРИ(1) + критерии(nCrit) + Итого(1) + метка суммы(1) + сумма(1)

  ws.getColumn(1).width = 26;
  for (let i = 0; i < nCrit; i++) ws.getColumn(2 + i).width = 12;
  ws.getColumn(2 + nCrit).width = 10;
  ws.getColumn(3 + nCrit).width = 12;
  ws.getColumn(4 + nCrit).width = 12;

  // группируем оценки по паре "участник + конкурсный номер", в порядке появления
  const groups = [];
  const byKey = new Map();
  store.submissions.forEach((s) => {
    const key = s.target + '|' + (s.targetGroup || '');
    if (!byKey.has(key)) { byKey.set(key, groups.length); groups.push({ target: s.target, targetGroup: s.targetGroup || '', items: [] }); }
    groups[byKey.get(key)].items.push(s);
  });

  let r = 1;

  if (groups.length === 0) {
    ws.getCell(1, 1).value = 'Пока нет ни одной оценки';
    ws.getCell(1, 1).font = { italic: true };
  }

  groups.forEach((g) => {
    const rows = g.items;

    // строка "УЧАСТНИК:"
    ws.mergeCells(r, 1, r, 2);
    ws.getCell(r, 1).value = 'УЧАСТНИК:';
    ws.getCell(r, 1).font = { bold: true };
    for (let c = 1; c <= 2; c++) ws.getCell(r, c).border = border;
    ws.mergeCells(r, 3, r, lastCol);
    ws.getCell(r, 3).value = g.target;
    ws.getCell(r, 3).font = { bold: true };
    for (let c = 3; c <= lastCol; c++) ws.getCell(r, c).border = border;
    r++;

    // строка "КОНКУРСНЫЙ НОМЕР:"
    ws.mergeCells(r, 1, r, 2);
    ws.getCell(r, 1).value = 'КОНКУРСНЫЙ НОМЕР:';
    ws.getCell(r, 1).font = { bold: true };
    for (let c = 1; c <= 2; c++) ws.getCell(r, c).border = border;
    ws.mergeCells(r, 3, r, lastCol);
    ws.getCell(r, 3).value = g.targetGroup || '—';
    ws.getCell(r, 3).font = { bold: true };
    for (let c = 3; c <= lastCol; c++) ws.getCell(r, c).border = border;
    r++;

    // шапка таблицы оценок
    ws.getCell(r, 1).value = 'ЧЛЕН ЖЮРИ';
    ws.getCell(r, 1).font = { bold: true };
    ws.getCell(r, 1).border = border;
    ws.mergeCells(r, 2, r, 1 + nCrit);
    ws.getCell(r, 2).value = 'Оценки';
    ws.getCell(r, 2).font = { bold: true };
    ws.getCell(r, 2).alignment = { horizontal: 'center' };
    for (let c = 2; c <= 1 + nCrit; c++) ws.getCell(r, c).border = border;
    ws.getCell(r, 2 + nCrit).value = 'Итого';
    ws.getCell(r, 2 + nCrit).font = { bold: true };
    ws.getCell(r, 2 + nCrit).border = border;
    r++;

    const firstJudgeRow = r;
    rows.forEach((s) => {
      ws.getCell(r, 1).value = s.expertName;
      ws.getCell(r, 1).border = border;
      s.scores.forEach((sc, ci) => {
        const cell = ws.getCell(r, 2 + ci);
        cell.value = sc.score;
        cell.alignment = { horizontal: 'center' };
        cell.font = { bold: true };
        cell.border = border;
      });
      const totalCell = ws.getCell(r, 2 + nCrit);
      totalCell.value = s.total;
      totalCell.alignment = { horizontal: 'center' };
      totalCell.font = { bold: true };
      totalCell.border = border;
      r++;
    });
    const lastJudgeRow = r - 1;

    // крупный блок "СУММА:" справа, объединённый по всем строкам жюри
    const grandTotal = rows.reduce((a, s) => a + s.total, 0);
    ws.mergeCells(firstJudgeRow, 3 + nCrit, lastJudgeRow, 3 + nCrit);
    const sumLbl = ws.getCell(firstJudgeRow, 3 + nCrit);
    sumLbl.value = 'СУММА:';
    sumLbl.font = { bold: true };
    sumLbl.alignment = { vertical: 'middle', horizontal: 'center' };
    sumLbl.border = border;
    ws.mergeCells(firstJudgeRow, 4 + nCrit, lastJudgeRow, 4 + nCrit);
    const sumVal = ws.getCell(firstJudgeRow, 4 + nCrit);
    sumVal.value = grandTotal;
    sumVal.font = { bold: true, size: 16 };
    sumVal.alignment = { vertical: 'middle', horizontal: 'center' };
    const medium = { style: 'medium', color: { argb: 'FF000000' } };
    sumVal.border = { top: medium, left: medium, bottom: medium, right: medium };

    r++;

    // заголовок "КОММЕНТАРИИ:"
    ws.mergeCells(r, 2, r, lastCol);
    ws.getCell(r, 1).value = '';
    ws.getCell(r, 2).value = 'КОММЕНТАРИИ:';
    ws.getCell(r, 2).font = { bold: true };
    r++;

    // комментарий каждого судьи отдельной строкой
    rows.forEach((s) => {
      ws.getCell(r, 1).value = s.expertName;
      ws.getCell(r, 1).font = { bold: true };
      ws.getCell(r, 1).border = border;
      ws.getCell(r, 1).alignment = { wrapText: true, vertical: 'top' };
      ws.mergeCells(r, 2, r, lastCol);
      const cCell = ws.getCell(r, 2);
      cCell.value = s.comment || '';
      cCell.alignment = { wrapText: true, vertical: 'top' };
      cCell.border = border;
      ws.getRow(r).height = 32;
      r++;
    });

    r += 2; // отступ перед следующим участником
  });

  await wb.xlsx.writeFile(XLSX_FILE);
}

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

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

io.on('connection', (socket) => {
  socket.emit('init', store);

  // Пароль модератора проверяется на сервере — в браузере он не хранится
  // и не виден через "просмотр кода страницы".
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
      target: typeof cfg.target === 'string' ? cfg.target.slice(0, 150) : store.config.target,
      targetGroup: typeof cfg.targetGroup === 'string' ? cfg.targetGroup.slice(0, 150) : store.config.targetGroup,
    };
    persist();
    io.emit('config:update', store.config);
    if (typeof cb === 'function') cb({ ok: true });
  });

  socket.on('submission:add', (sub) => {
    if (!sub || !sub.expertName || !Array.isArray(sub.scores)) return;
    if (sub.scores.length !== store.config.criteria.length) return;
    if (String(sub.comment || '').trim().length === 0) return;

    const clean = {
      id: Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      expertId: Number.isInteger(sub.expertId) ? sub.expertId : null,
      expertName: String(sub.expertName).slice(0, 100),
      target: store.config.target && store.config.target.trim() ? store.config.target.trim() : '—',
      targetGroup: store.config.targetGroup && store.config.targetGroup.trim() ? store.config.targetGroup.trim() : '',
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
