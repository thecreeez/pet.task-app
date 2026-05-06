const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, BorderStyle, ShadingType } = require('docx');

const app = express();
const PORT = 2048;
const DATA_FILE = path.join(__dirname, 'data', 'data.json');

fs.mkdirSync(path.join(__dirname, 'data'), { recursive: true });
fs.mkdirSync(path.join(__dirname, 'public', 'uploads'), { recursive: true });

if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify({ modules: [], tasks: [], steps: [] }, null, 2));
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'public', 'uploads')));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, path.join(__dirname, 'public', 'uploads')),
  filename: (req, file, cb) => cb(null, uuidv4() + path.extname(file.originalname))
});
const upload = multer({ storage });

function readData() { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
function writeData(data) { fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2)); }

app.get('/api/data', (req, res) => res.json(readData()));
app.post('/api/data', (req, res) => { writeData(req.body); res.json({ ok: true }); });
app.post('/api/upload', upload.single('image'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file' });
  res.json({ url: '/uploads/' + req.file.filename });
});

// ── Format helpers ─────────────────────────────────────────────────────────
function fmt(d) {
  return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${String(d.getFullYear()).slice(2)}`;
}
function fmtLong(d) {
  return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`;
}

// ── Fetch a single day from the kuzyak.in per-day endpoint ────────────────
// Response shape: { isWorkingDay: bool, isShortDay: bool, holiday: string|null, ... }
async function fetchDay(date) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  const day = date.getDate();
  try {
    const r = await fetch(`https://calendar.kuzyak.in/api/calendar/${y}/${m}/${day}`);
    if (!r.ok) return null;
    return await r.json(); // { isWorkingDay, isShortDay, holiday, ... }
  } catch {
    return null;
  }
}

// ── Core: fetch per-day data and compute smart working range ───────────────
// Each day in returned `days` array has:
//   { date, isWorking, isShortDay, holiday, dayName, dateStr }
// rangeStr / rangeStrLong are trimmed to first→last working day of the week.
async function computeWeekRange(offsetWeeks) {
  const now = new Date();
  const dow = now.getDay();
  const diffToMon = dow === 0 ? -6 : 1 - dow;
  const mon = new Date(now);
  mon.setDate(now.getDate() + diffToMon + offsetWeeks * 7);
  mon.setHours(0, 0, 0, 0);

  // Mon–Fri Date objects
  const weekDays = Array.from({ length: 5 }, (_, i) => {
    const d = new Date(mon);
    d.setDate(mon.getDate() + i);
    return d;
  });

  // Fetch all 5 days in parallel
  const apiResults = await Promise.all(weekDays.map(fetchDay));

  // Annotate each day with clean fields
  const days = weekDays.map((d, i) => {
    const api = apiResults[i];
    // If API returned data, trust isWorkingDay; otherwise assume working
    const isWorking  = api ? api.isWorkingDay  : true;
    const isShortDay = api ? (api.isShortDay || false) : false;
    const holiday    = api ? (api.holiday || null) : null;

    return {
      date:      d,
      dateStr:   fmtLong(d),
      dayName:   ['Вс','Пн','Вт','Ср','Чт','Пт','Сб'][d.getDay()],
      isWorking,
      isShortDay,
      holiday,   // e.g. "Праздник Весны и Труда" or null
    };
  });

  const working = days.filter(d => d.isWorking);
  const firstW  = working[0]                      || days[0];
  const lastW   = working[working.length - 1]     || days[4];

  return {
    rangeStr:        `${fmt(firstW.date)}-${fmt(lastW.date)}`,
    rangeStrLong:    `${fmtLong(firstW.date)} – ${fmtLong(lastW.date)}`,
    workingCount:    working.length,
    nonWorkingCount: days.length - working.length,
    days,
  };
}

// ── Proxy: expose week-range data to the frontend ──────────────────────────
app.get('/api/week-range/:offset', async (req, res) => {
  try {
    const info = await computeWeekRange(parseInt(req.params.offset) || 0);
    // Return serialisable version (no Date objects)
    res.json({
      ...info,
      days: info.days.map(d => ({
        dateStr:   d.dateStr,
        dayName:   d.dayName,
        isWorking: d.isWorking,
        isShortDay: d.isShortDay,
        holiday:   d.holiday,   // string like "Праздник Весны и Труда" or null
      }))
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Export ─────────────────────────────────────────────────────────────────
app.post('/api/export', async (req, res) => {
  try {
    const { tasks, modules, employee, periodOffset = 0 } = req.body;
    const readyTasks    = tasks.filter(t => t.column === 'ready');
    const nextWeekTasks = tasks.filter(t => t.column === 'next-week');
    const getModuleName = id => (modules.find(m => m.id === id) || {}).name || '';

    // Compute smart ranges in parallel
    const [curWeek, nxtWeek] = await Promise.all([
      computeWeekRange(periodOffset),
      computeWeekRange(periodOffset + 1),
    ]);

    const font = 'Times New Roman';
    const CONTENT_W = 9355;
    const COL1 = 1500;
    const COL2 = CONTENT_W - COL1;
    const cm = { top: 80, bottom: 80, left: 120, right: 120 };
    const sb = { style: BorderStyle.SINGLE, size: 4, color: '999999' };
    const nb = { style: BorderStyle.NIL };
    const allB  = { top: sb, bottom: sb, left: sb, right: sb };
    const noTop = { top: nb, bottom: sb, left: sb, right: sb };
    const green = { fill: 'E2EFD9', type: ShadingType.CLEAR };

    function makeTasksParagraphs(taskList) {
      if (!taskList.length) return [new Paragraph({ children: [] })];
      return taskList.map(task => {
        const mod = getModuleName(task.module);
        const children = [new TextRun({ text: task.title, font, size: 26 })];
        if (mod) children.unshift(new TextRun({ text: mod + ' - ', font, size: 26, bold: true }));
        return new Paragraph({ children });
      });
    }

    const reportTable = new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [COL1, COL2],
      rows: [
        new TableRow({ children: [
          new TableCell({ borders: allB, shading: green, width: { size: COL1, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Дата', bold: true, font, size: 28 })] })] }),
          new TableCell({ borders: allB, shading: green, width: { size: COL2, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Выполненные задачи (% готовности)', bold: true, font, size: 28 })] })] }),
        ]}),
        new TableRow({ children: [
          new TableCell({ borders: allB, width: { size: COL1, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: curWeek.rangeStr, bold: true, font, size: 26 })] })] }),
          new TableCell({ borders: allB, width: { size: COL2, type: WidthType.DXA }, margins: cm,
            children: makeTasksParagraphs(readyTasks) }),
        ]}),
        new TableRow({ children: [
          new TableCell({ borders: allB, shading: green, width: { size: CONTENT_W, type: WidthType.DXA }, columnSpan: 2, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'План на следующую неделю', bold: true, font, size: 28, color: '000000' })] })] }),
        ]}),
        new TableRow({ children: [
          new TableCell({ borders: noTop, width: { size: COL1, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: nxtWeek.rangeStr, font, size: 24 })] })] }),
          new TableCell({ borders: noTop, width: { size: COL2, type: WidthType.DXA }, margins: cm,
            children: makeTasksParagraphs(nextWeekTasks) }),
        ]}),
      ]
    });

    const doc = new Document({
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Недельный отчет о проделанной работе по дням', bold: true, size: 36, font })] }),
          new Paragraph({ children: [new TextRun({ text: 'Период описанной работы: ', bold: true, size: 28, font }), new TextRun({ text: curWeek.rangeStr, bold: true, size: 26, font })] }),
          new Paragraph({ children: [new TextRun({ text: 'Сотрудник: ', bold: true, size: 28, font }), new TextRun({ text: employee || '', size: 28, font })] }),
          reportTable,
          new Paragraph({ children: [] }),
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Disposition', `attachment; filename="report_${Date.now()}.docx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Report app running on http://localhost:${PORT}`));