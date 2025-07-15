const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const Jimp = require('jimp');

const app = express();
app.use(cors({ origin: '*' })); // Cambia '*' por tu dominio de frontend en producción
app.use(express.json());

const upload = multer({ dest: 'fotos_temp/' });
if (!fs.existsSync('fotos_temp')) fs.mkdirSync('fotos_temp');

// Gestión de trabajos en memoria
const jobs = {};

// Endpoint: grupos disponibles
app.get('/api/grupos', (req, res) => {
  try {
    const workbook = XLSX.readFile('./grupos.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheet);
    const nombres = [...new Set(grupos.map(row => row.grupo).filter(Boolean))].sort();
    res.json({ grupos: nombres });
  } catch (err) {
    res.status(500).json({ error: "No se pudieron obtener los grupos." });
  }
});

// Inicia generación del Excel
app.post('/api/genera-excel-final-async', async (req, res) => {
  try {
    const { grupo, fotosDisponibles = [], origenFotos = "local", sinImagenes = false } = req.body;
    const jobId = uuidv4();
    jobs[jobId] = { progress: 0, buffer: null, error: null, filename: null, startedAt: Date.now(), fotosLocales: {} };
    // Si usa modo local, devuelve la lista de imágenes requeridas
    if (origenFotos === "local" && !sinImagenes) {
      const workbook = XLSX.readFile('./grupos.xlsx');
      const sh = workbook.Sheets[workbook.SheetNames[0]];
      const grupos = XLSX.utils.sheet_to_json(sh);
      const codigos = grupos.filter(row => row.grupo === grupo).map(row => row.codigo?.toString());
      const requiredFotos = codigos.map(c => [`${c}.jpg`, `${c}.jpeg`, `${c}.png`]).flat();
      const faltan = requiredFotos.filter(f => !fotosDisponibles.includes(f.toLowerCase()));
      return res.json({ jobId, requiredFotos: faltan });
    }
    // Si no, dispara generación directa
    generarExcelAsync(req.body, jobId);
    res.json({ jobId });
  } catch (err) {
    res.status(500).json({ error: "Error iniciando la generación del Excel." });
  }
});

// Subida de fotos temporales
app.post('/api/subir-fotos/:jobId', upload.array('fotos'), (req, res) => {
  const { jobId } = req.params;
  if (!jobs[jobId]) return res.status(404).json({ error: 'Job no encontrado' });
  req.files.forEach(f => {
    jobs[jobId].fotosLocales[f.originalname.toLowerCase()] = f.path;
  });
  res.json({ ok: true });
});

// Progreso del trabajo
app.get('/api/progreso/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job) return res.status(404).json({ error: 'Trabajo no encontrado' });
  let eta = null;
  if (job.progress > 2 && job.progress < 99 && job.startedAt) {
    const elapsed = (Date.now() - job.startedAt) / 1000;
    const p = Math.max(job.progress, 1) / 100;
    const total = elapsed / p;
    eta = Math.max(0, Math.round(total - elapsed));
  }
  res.json({ progress: job.progress, error: job.error, filename: job.filename, eta });
});

// Descarga del Excel
app.get('/api/descarga-excel/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job || !job.buffer) return res.status(404).json({ error: 'Archivo no disponible.' });
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(job.buffer);
});

// Generador Excel adaptado, busca primero imágenes locales.
async function generarExcelAsync(params, jobId) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false, origenFotos = "local" } = params;
    const workbookGrupos = XLSX.readFile('./grupos.xlsx');
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    const codigosGrupo = grupos.filter(row => row.grupo === grupo).map(row => row.codigo?.toString());
    if (!codigosGrupo.length) {
      jobs[jobId].error = "No hay artículos para ese grupo.";
      jobs[jobId].progress = 100;
      return;
    }
    // Simula consulta API/BD (implementa aquí tu lógica real de artículos)
    let articulos = codigosGrupo.map((codigo, i) => ({
      codigo, descripcion: `Demo Artículo ${codigo}`, disponible: 10 + i, precioVenta: 9.95 + i, ean13: "0000", umv: 1 // demo data
    }));
    // Filtra stock si corresponde
    if (soloStock) articulos = articulos.filter(a => a.disponible > 0);
    // Crea Excel
    const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Listado');
    ws.addRow(campos);
    for (let [i, art] of articulos.entries()) {
      let fila = campos.map(c => c === "imagen" ? "" : art[c] ?? "");
      ws.addRow(fila);
      // Imagen local
      if (!sinImagenes && origenFotos === "local" && jobs[jobId].fotosLocales) {
        for (let ext of ['jpg', 'jpeg', 'png']) {
          const nombre = `${art.codigo}.${ext}`.toLowerCase();
          if (jobs[jobId].fotosLocales[nombre]) {
            const buffer = fs.readFileSync(jobs[jobId].fotosLocales[nombre]);
            let img = await Jimp.read(buffer);
            img.cover(110, 110); // Tamaño fijo para Excel
            const miniBuffer = await img.getBufferAsync(Jimp.MIME_JPEG);
            const imageId = workbook.addImage({ buffer: miniBuffer, extension: 'jpeg' });
            ws.addImage(imageId, { tl: { col: campos.length - 1, row: i + 1 }, ext: { width: 110, height: 110 } });
            break;
          }
        }
      }
      jobs[jobId].progress = Math.round(((i + 1) / articulos.length) * 90);
    }
    let buffer = await workbook.xlsx.writeBuffer();
    jobs[jobId].buffer = Buffer.from(buffer);
    jobs[jobId].progress = 100;
    jobs[jobId].filename = `listado_${grupo}_${idioma}${sinImagenes ? '_sinImagenes' : ''}.xlsx`;
    // Borra imágenes temporales
    if (jobs[jobId].fotosLocales) {
      Object.values(jobs[jobId].fotosLocales).forEach(f => { try { fs.unlinkSync(f); } catch { } });
      jobs[jobId].fotosLocales = {};
    }
  } catch (err) {
    jobs[jobId].error = "Error generando el Excel.";
    jobs[jobId].progress = 100;
  }
}

app.listen(3000, () => console.log('Servidor escuchando en puerto 3000'));
