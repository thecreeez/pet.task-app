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

app.post('/api/export', async (req, res) => {
  try {
    const { tasks, modules, employee } = req.body;
    const readyTasks    = tasks.filter(t => t.column === 'ready');
    const nextWeekTasks = tasks.filter(t => t.column === 'next-week');
    const getModuleName = id => (modules.find(m => m.id === id) || {}).name || '';

    function getWeekRange(offsetWeeks) {
      const now = new Date();
      const day = now.getDay();
      const diffToMon = day === 0 ? -6 : 1 - day;
      const mon = new Date(now);
      mon.setDate(now.getDate() + diffToMon + offsetWeeks * 7);
      const fri = new Date(mon);
      fri.setDate(mon.getDate() + 4);
      const fmt = d => `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${String(d.getFullYear()).slice(2)}`;
      return `${fmt(mon)}-${fmt(fri)}`;
    }

    const currentWeekRange = getWeekRange(0);
    const nextWeekRange    = getWeekRange(1);
    const font = 'Times New Roman';
    const COL1 = 1463, COL2 = 8975, TOTAL = COL1 + COL2;
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
        let children = [new TextRun({ text: task.title, font, size: 26 })];
        if (mod) {
          children.unshift(new TextRun({ text: mod + " - ", font, size: 26, bold: true }))
        }
        return new Paragraph({ children });
      });
    }

    const reportTable = new Table({
      width: { size: TOTAL, type: WidthType.DXA },
      columnWidths: [COL1, COL2],
      rows: [
        // Row 1: headers
        new TableRow({ children: [
          new TableCell({ borders: allB, shading: green, width: { size: COL1, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Дата', bold: true, font, size: 28 })] })] }),
          new TableCell({ borders: allB, shading: green, width: { size: COL2, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Выполненные задачи (% готовности)', bold: true, font, size: 28 })] })] }),
        ]}),
        // Row 2: current week date | ready tasks
        new TableRow({ children: [
          new TableCell({ borders: allB, width: { size: COL1, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: currentWeekRange, bold: true, font, size: 26 })] })] }),
          new TableCell({ borders: allB, width: { size: COL2, type: WidthType.DXA }, margins: cm,
            children: makeTasksParagraphs(readyTasks) }),
        ]}),
        // Row 3: "План на следующую неделю" — merged full-width green header
        new TableRow({ children: [
          new TableCell({ borders: allB, shading: green, width: { size: TOTAL, type: WidthType.DXA }, columnSpan: 2, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'План на следующую неделю', bold: true, font, size: 28, color: '000000' })] })] }),
        ]}),
        // Row 4: next week date | next-week tasks (no top border like template)
        new TableRow({ children: [
          new TableCell({ borders: noTop, width: { size: COL1, type: WidthType.DXA }, margins: cm,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: nextWeekRange, font, size: 24 })] })] }),
          new TableCell({ borders: noTop, width: { size: COL2, type: WidthType.DXA }, margins: cm,
            children: makeTasksParagraphs(nextWeekTasks) }),
        ]}),
      ]
    });

    const doc = new Document({
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Недельный отчет о проделанной работе по дням', bold: true, size: 36, font })] }),
          new Paragraph({ children: [new TextRun({ text: 'Период описанной работы: ', bold: true, size: 28, font }), new TextRun({ text: currentWeekRange, bold: true, size: 26, font })] }),
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
