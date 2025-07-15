const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const https = require('https');
const { v4: uuidv4 } = require('uuid');
const Jimp = require('jimp');
const pLimit = require('p-limit');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors({
  origin: 'https://webb2b.netlify.app'
}));
app.use(express.json());

const upload = multer({ dest: 'fotos_temp/' });

if (!fs.existsSync('fotos_temp')) fs.mkdirSync('fotos_temp');

const diccionario_traduccion = {
  Español: { codigo: "Código", descripcion: "Descripción", disponible: "Disponible", ean13: "EAN13", precioVenta: "Precio", umv: "UMV", imagen: "Imagen" },
  Inglés: { codigo: "Code", descripcion: "Description", disponible: "Available", ean13: "EAN13", precioVenta: "Price", umv: "MOQ", imagen: "Image" },
  Francés: { codigo: "Code", descripcion: "Description", disponible: "Disponible", ean13: "EAN13", precioVenta: "Prix", umv: "MOQ", imagen: "Image" },
  Italiano: { codigo: "Codice", descripcion: "Descrizione", disponible: "Disponibile", ean13: "EAN13", precioVenta: "Prezzo", umv: "MOQ", imagen: "Immagine" }
};

const usuarios_api = {
  Español: { usuario: "amazon@espana.es", password: "0glLD6g7Dg" },
  Inglés: { usuario: "ingles@atosa.es", password: "AtosaIngles" },
  Francés: { usuario: "frances@atosa.es", password: "AtosaFrancés" },
  Italiano: { usuario: "italiano@atosa.es", password: "AtosaItaliano" }
};

// Endpoint para grupos usando archivo local
app.get('/api/grupos', async (req, res) => {
  try {
    const workbookGrupos = XLSX.readFile('./grupos.xlsx');
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    const nombres = [...new Set(grupos.map(row => row.grupo).filter(Boolean))].sort();
    res.json({ grupos: nombres });
  } catch (err) {
    console.error("[/api/grupos] Error:", err);
    res.status(500).json({ error: "No se pudieron obtener los grupos." });
  }
});

// ---- GESTIÓN DE JOBS Y PROGRESO EN MEMORIA ----
const jobs = {}; // jobId: { progress, buffer, error, filename, startedAt, fotosLocales }

// 1. Inicia generación de Excel
app.post('/api/genera-excel-final-async', async (req, res) => {
  try {
    const { grupo, fotosDisponibles = [], origenFotos = "local", sinImagenes = false } = req.body;
    const jobId = uuidv4();
    jobs[jobId] = { progress: 0, buffer: null, error: null, filename: null, startedAt: Date.now(), fotosLocales: {} };

    // Si se selecciona "local", devolvemos la lista de imágenes requeridas (que el frontend debe subir)
    if (origenFotos === "local" && !sinImagenes) {
      let workbookGrupos = XLSX.readFile('./grupos.xlsx');
      let sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
      let grupos = XLSX.utils.sheet_to_json(sheetGrupos);
      let codigosGrupo = grupos.filter(row => row.grupo === grupo).map(row => row.codigo?.toString());
      const requiredFotos = codigosGrupo.map(c => [`${c}.jpg`, `${c}.jpeg`, `${c}.png`]).flat();
      const faltan = requiredFotos.filter(f => !fotosDisponibles.includes(f.toLowerCase()));
      return res.json({ jobId, requiredFotos: faltan });
    }

    generarExcelAsync(req.body, jobId);
    res.json({ jobId });
  } catch (err) {
    console.error("[/api/genera-excel-final-async] Error:", err);
    res.status(500).json({ error: "Error iniciando la generación del Excel." });
  }
});

// 2. Subida de fotos locales (solo cuando origenFotos = "local")
app.post('/api/subir-fotos/:jobId', upload.array('fotos'), (req, res) => {
  const { jobId } = req.params;
  if (!jobs[jobId]) return res.status(404).json({ error: 'Job no encontrado' });
  req.files.forEach(f => {
    jobs[jobId].fotosLocales[f.originalname.toLowerCase()] = f.path;
  });
  res.json({ ok: true });
});

// 3. Progreso
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

// 4. Descarga el archivo generado
app.get('/api/descarga-excel/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job || !job.buffer) return res.status(404).json({ error: 'Archivo no disponible.' });
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(job.buffer);
});

// 5. Generador Excel
async function generarExcelAsync(params, jobId) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false, origenFotos = "local" } = params;
    const maxFilas = 3500;

    // Leer grupos.xlsx local
    let workbookGrupos = XLSX.readFile('./grupos.xlsx');
    let sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    let grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    let codigosGrupo = grupos.filter(row => row.grupo === grupo).map(row => row.codigo?.toString());
    if (!codigosGrupo.length) {
      jobs[jobId].error = "No hay artículos para ese grupo."; jobs[jobId].progress = 100; return;
    }

    // Llama a la API de Atosa
    const { usuario, password } = usuarios_api[idioma] || usuarios_api["Español"];
    const apiURL = "https://b2b.atosa.es:880/api/articulos/";
    let respArticulos = await axios.get(apiURL, {
      auth: { username: usuario, password: password },
      timeout: 60000,
      httpsAgent: new https.Agent({ rejectUnauthorized: false }),
    });
    let articulos = respArticulos.data.filter(art =>
      codigosGrupo.includes(art.codigo?.toString()) &&
      (!soloStock || parseInt(art.disponible || 0) > 0)
    );
    articulos = articulos.slice(0, maxFilas);
    if (!articulos.length) {
      jobs[jobId].error = "No hay artículos que coincidan con el filtro."; jobs[jobId].progress = 100; return;
    }

    // Descuentos (como antes)
    let articulos_sin_descuento = new Set();
    if (descuento > 0) {
      try {
        const usuariosDescuento = [
          { usuario: "compras@b2cmarketonline.es", password: "rXCRzzWKI6" },
          { usuario: "santi@tradeinn.com", password: "C8Zg1wqgfe" }
        ];
        const [resp4, resp8] = await Promise.all(usuariosDescuento.map(u =>
          axios.get(apiURL, {
            auth: { username: u.usuario, password: u.password },
            timeout: 60000,
            httpsAgent: new https.Agent({ rejectUnauthorized: false }),
          })
        ));
        const precios4 = Object.fromEntries(resp4.data.map(a => [a.codigo, parseFloat(a.precioVenta)]));
        const precios8 = Object.fromEntries(resp8.data.map(a => [a.codigo, parseFloat(a.precioVenta)]));
        articulos.forEach(art => {
          const cod = art.codigo;
          const pv0 = parseFloat(art.precioVenta);
          const pv4 = precios4[cod];
          const pv8 = precios8[cod];
          if (
            pv4 !== undefined && pv8 !== undefined &&
            Math.abs(pv0 - pv4) < 0.01 && Math.abs(pv0 - pv8) < 0.01
          ) {
            articulos_sin_descuento.add(cod);
          }
        });
      } catch (err) { }
    }

    // Crear Excel
    const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
    const traducido = campos.map(c => diccionario_traduccion[idioma][c]);
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Listado');
    ws.addRow(traducido);
    const colWidths = [12, 40, 12, 12, 12, 10, 18];
    ws.columns = ws.columns.map((col, idx) => ({ ...col, width: colWidths[idx] || 15 }));
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    if (sinImagenes) {
      for (let [i, art] of articulos.entries()) {
        let fila = [];
        for (let c of campos) {
          let valor = art[c] ?? "";
          if (c === "precioVenta") {
            try {
              if (descuento > 0 && !articulos_sin_descuento.has(art.codigo)) {
                valor = Math.round((parseFloat(valor) * (1 - descuento / 100)) * 100) / 100;
              } else {
                valor = parseFloat(valor);
              }
            } catch { }
          }
          fila.push(valor);
        }
        ws.addRow(fila);
        jobs[jobId].progress = Math.round(((i + 1) / articulos.length) * 80);
      }
    } else {
      const limit = pLimit(8);

      for (let [i, art] of articulos.entries()) {
        let fila = [];
        for (let c of campos) {
          let valor = art[c] ?? "";
          if (c === "precioVenta") {
            try {
              if (descuento > 0 && !articulos_sin_descuento.has(art.codigo)) {
                valor = Math.round((parseFloat(valor) * (1 - descuento / 100)) * 100) / 100;
              } else {
                valor = parseFloat(valor);
              }
            } catch { }
          }
          fila.push(valor);
        }
        ws.addRow(fila);
      }

      await Promise.all(articulos.map((art, i) => limit(async () => {
        let fotoBuffer = null;
        if (origenFotos === "local" && jobs[jobId].fotosLocales) {
          for (let ext of ['jpg', 'jpeg', 'png']) {
            const nombre = `${art.codigo}.${ext}`.toLowerCase();
            if (jobs[jobId].fotosLocales[nombre]) {
              fotoBuffer = fs.readFileSync(jobs[jobId].fotosLocales[nombre]);
              break;
            }
          }
        } else if (origenFotos === "api") {
          fotoBuffer = await obtenerFotoArticuloAPI(art.codigo);
        }
        if (fotoBuffer) {
          try {
            let img = await Jimp.read(fotoBuffer);
            img.cover(110, 110);
            img.quality(60);
            const miniBuffer = await img.getBufferAsync(Jimp.MIME_JPEG);
            const imageId = workbook.addImage({
              buffer: miniBuffer,
              extension: 'jpeg'
            });
            ws.addImage(imageId, {
              tl: { col: campos.length - 1, row: i + 1 },
              ext: { width: 110, height: 110 }
            });
          } catch (e) {}
        }
        jobs[jobId].progress = Math.round(((i + 1) / articulos.length) * 80);
      })));
    }

    ws.eachRow({ includeEmpty: false }, function(row, rowNumber) {
      row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      row.height = 90;
      if (rowNumber === 1) row.font = { bold: true, size: 14 };
      else row.font = { size: 13 };
    });

    let buffer = await workbook.xlsx.writeBuffer();
    jobs[jobId].buffer = Buffer.from(buffer);
    jobs[jobId].progress = 100;
    jobs[jobId].filename = `listado_${grupo}_${idioma}${sinImagenes ? '_sinImagenes' : ''}.xlsx`;

    // Limpia las fotos temporales
    if (jobs[jobId].fotosLocales) {
      Object.values(jobs[jobId].fotosLocales).forEach(f => {
        try { fs.unlinkSync(f); } catch { }
      });
      jobs[jobId].fotosLocales = {};
    }
  } catch (err) {
    jobs[jobId].error = "Error generando el Excel (excepción interna).";
    jobs[jobId].progress = 100;
    console.error("[generarExcelAsync] Error:", err);
  }
}

async function obtenerFotoArticuloAPI(codigo) {
  const usuario = usuarios_api["Español"].usuario;
  const password = usuarios_api["Español"].password;
  try {
    const fotoResp = await axios.get(
      `https://b2b.atosa.es:880/api/articulos/foto/${codigo}`,
      {
        auth: { username: usuario, password: password },
        timeout: 10000,
        httpsAgent: new https.Agent({ rejectUnauthorized: false }),
      }
    );
    const fotos = fotoResp.data.fotos;
    if (Array.isArray(fotos) && fotos.length > 0) {
      return Buffer.from(fotos[0], 'base64');
    }
    return null;
  } catch (e) {
    return null;
  }
}

app.get('/', (req, res) => {
  res.send('Servidor ATOSA backend funcionando.');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
});