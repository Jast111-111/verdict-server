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
  config: { eventName: 'Мероприятие', criteria: [...DEFAULT_CRITERIA], experts: [...DEFAULT_EXPERTS] },
  submissions: [],
};

// на случай, если store.json был сохранён старой версией сервера — дополняем недостающее
if (!Array.isArray(store.config.experts) || store.config.experts.length !== 4) {
  store.config.experts = [...DEFAULT_EXPERTS];
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
  const ws = wb.addWorksheet('Оценки');

  const cols = [
    { header: '#', key: 'n', width: 5 },
    { header: 'Время', key: 'time', width: 18 },
    { header: 'Эксперт', key: 'expert', width: 20 },
    { header: 'Оцениваемый', key: 'target', width: 24 },
  ];
  store.config.criteria.forEach((c, i) => {
    cols.push({ header: c, key: 'score' + i, width: 12 });
  });
  cols.push({ header: 'Итого', key: 'total', width: 10 });
  cols.push({ header: 'Средний балл', key: 'avg', width: 14 });
  cols.push({ header: 'Комментарий', key: 'comment', width: 40 });
  ws.columns = cols;
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFE7CE' } };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  store.submissions.forEach((s, i) => {
    const row = {
      n: i + 1,
      time: new Date(s.timestamp).toLocaleString('ru-RU'),
      expert: s.expertName,
      target: s.target,
      total: s.total,
      avg: s.average,
      comment: s.comment || '',
    };
    s.scores.forEach((sc, idx) => {
      row['score' + idx] = sc.score;
    });
    ws.addRow(row);
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
      experts: Array.isArray(cfg.experts) && cfg.experts.length === 4
        ? cfg.experts.map((e) => String(e || '').slice(0, 100))
        : store.config.experts,
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
      target: String(sub.target || '—').slice(0, 150),
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
