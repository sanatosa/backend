// server.js — ATOSA Excel + histórico de altas/bajas de artículos

const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const https = require('https');
const { v4: uuidv4 } = require('uuid');
const Jimp = require('jimp');
const pLimit = require('p-limit').default;
require('dotenv').config();
const nodemailer = require('nodemailer');
const fs = require('fs');

const app = express();

// --- Configuración CORS para Netlify frontend ---
app.use(cors({
  origin: 'https://webb2b.netlify.app',
  credentials: true,
  methods: ['GET', 'POST', 'OPTIONS']
}));
app.options('*', cors());
app.use(express.json());

// --- Config fijar alto de filas ---
const imagenPx = 110;
const filaAltura = 82.0; // Altura de fila Excel

const diccionario_traduccion = {
  Español: {
    codigo: "Código", descripcion: "Descripción", disponible: "Stock",
    ean13: "EAN", precioVenta: "Precio", umv: "UMV", imagen: "Imagen"
  },
  Inglés: {
    codigo: "Code", descripcion: "Description", disponible: "Available",
    ean13: "EAN", precioVenta: "Price", umv: "MOQ", imagen: "Image"
  },
  Francés: {
    codigo: "Code", descripcion: "Description", disponible: "Stock",
    ean13: "EAN", precioVenta: "Prix", umv: "MOQ", imagen: "Image"
  },
  Italiano: {
    codigo: "Codice", descripcion: "Descrizione", disponible: "Stock",
    ean13: "EAN", precioVenta: "Prezzo", umv: "MOQ", imagen: "Immagine"
  }
};

const usuarios_api = {
  Español: { usuario: "amazon@espana.es", password: "0glLD6g7Dg" },
  Inglés: { usuario: "ingles@atosa.es", password: "AtosaIngles" },
  Francés: { usuario: "frances@atosa.es", password: "AtosaFrances" },
  Italiano: { usuario: "italiano@atosa.es", password: "AtosaItaliano" }
};
const usuario8 = { usuario: "santi@tradeinn.com", password: "C8Zg1wqgfe" };

const jobs = {};

// --- Artículos históricos ---
const HISTORICO_PATH = './historico_articulos.json';

function loadHistorico() {
  if (!fs.existsSync(HISTORICO_PATH)) {
    return { fecha: null, codigos: [] };
  }
  try {
    const data = fs.readFileSync(HISTORICO_PATH, 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    console.error('Error cargando histórico:', e);
    return { fecha: null, codigos: [] };
  }
}
function saveHistorico(codigos) {
  const data = {
    fecha: new Date().toISOString(),
    codigos: codigos
  };
  fs.writeFileSync(HISTORICO_PATH, JSON.stringify(data, null, 2));
}
function compararListas(antes, ahora) {
  const setAntes = new Set(antes);
  const setAhora = new Set(ahora);
  const altas = ahora.filter(c => !setAntes.has(c));
  const bajas = antes.filter(c => !setAhora.has(c));
  return { altas, bajas };
}
function backupHistorico() {
  if (fs.existsSync(HISTORICO_PATH)) {
    fs.copyFileSync(HISTORICO_PATH, './historico_anterior.json');
  }
}

// --- Nueva funcionalidad: Cargar orden de artículos ---
let ordenArticulos = {};
function cargarOrdenArticulos() {
  try {
    const workbookOrden = XLSX.readFile('./orden.xlsx');
    const sheetOrden = workbookOrden.Sheets[workbookOrden.SheetNames[0]];
    const datosOrden = XLSX.utils.sheet_to_json(sheetOrden, { header: ['orden', 'codigo'] });
    ordenArticulos = {};
    datosOrden.forEach(row => {
      if (row.codigo && row.orden !== undefined) {
        ordenArticulos[row.codigo.toString().trim()] = parseInt(row.orden) || 999999;
      }
    });
    console.log(`Cargados ${Object.keys(ordenArticulos).length} artículos del archivo orden.xlsx`);
  } catch (error) {
    console.error('Error cargando orden.xlsx:', error.message);
    ordenArticulos = {};
  }
}
function ordenarArticulos(articulos) {
  return articulos.sort((a, b) => {
    const codigoA = a.codigo ? a.codigo.toString().trim() : '';
    const codigoB = b.codigo ? b.codigo.toString().trim() : '';
    const ordenA = ordenArticulos[codigoA] || 999999;
    const ordenB = ordenArticulos[codigoB] || 999999;
    if (ordenA === ordenB) {
      return codigoA.localeCompare(codigoB);
    }
    return ordenA - ordenB;
  });
}
// --- Al iniciar el servidor carga el orden del excel ---
cargarOrdenArticulos();

async function obtenerFotoArticuloAPI(codigo, usuario, password, intentos = 3) {
  for (let i = 0; i < intentos; i++) {
    try {
      const resp = await axios.get(`https://b2b.atosa.es:880/api/articulos/foto/${codigo}`, {
        auth: { username: usuario, password },
        timeout: 15000,
        httpsAgent: new https.Agent({ rejectUnauthorized: false }),
      });
      const fotos = resp.data.fotos;
      if (Array.isArray(fotos) && fotos.length > 0) {
        const buffer = Buffer.from(fotos[0], 'base64');
        if (buffer.length > 0) return buffer;
      }
    } catch (e) {
      console.log(`Intento ${i + 1} fallido para imagen ${codigo}:`, e.message);
      if (i < intentos - 1) await new Promise(r => setTimeout(r, 1000 * (i + 1)));
    }
  }
  return null;
}
function validarBuffer(buffer) {
  if (!buffer || buffer.length === 0) return false;
  const jpegHeader = buffer.slice(0, 2).toString('hex') === 'ffd8';
  const pngHeader = buffer.slice(0, 8).toString('hex') === '89504e470d0a1a0a';
  return jpegHeader || pngHeader;
}
async function crearImagenPorDefecto() {
  const img = new Jimp(imagenPx, imagenPx, '#f0f0f0');
  const font = await Jimp.loadFont(Jimp.FONT_SANS_16_BLACK);
  img.print(font, 10, imagenPx / 2 - 10, 'Sin imagen');
  return await img.getBufferAsync(Jimp.MIME_JPEG);
}
async function enviarEmailConAdjunto(emailDestino, bufferExcel, filename) {
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });
  const mailOptions = {
    from: process.env.EMAIL_FROM,
    to: emailDestino,
    subject: 'Tu archivo Excel está listo',
    text: 'Adjuntamos el listado generado. ¡Gracias por usar la herramienta!',
    attachments: [
      { filename, content: bufferExcel }
    ]
  };
  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email enviado a ${emailDestino}`);
  } catch (error) {
    console.error(`Error enviando email: ${error.message}`);
  }
}

// ------------------ ENDPOINTS ---------------------------

// Grupos disponibles
app.get('/api/grupos', async (req, res) => {
  try {
    const workbook = XLSX.readFile('./grupos.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheet);
    const nombres = [...new Set(grupos.map(row => (row.grupo ? row.grupo.toString().trim() : null)).filter(gr => gr && gr.length > 0))].sort();
    res.json({ grupos: nombres });
  } catch (err) {
    res.status(500).json({ error: "No se pudieron obtener los grupos." });
  }
});

// Excel asincrónico (principal)
app.post('/api/genera-excel-final-async', async (req, res) => {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false, email } = req.body;
    const jobId = uuidv4();
    jobs[jobId] = { progress: 0, buffer: null, error: null, filename: null, startedAt: Date.now(), fase: "Preparando" };
    generarExcelAsync({ grupo, idioma, descuento, soloStock, sinImagenes, email }, jobId);
    res.json({ jobId });
  } catch (err) {
    res.status(500).json({ error: "Error iniciando la generación del Excel." });
  }
});

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
  res.json({ progress: job.progress, error: job.error, filename: job.filename, eta, fase: job.fase });
});

app.get('/api/descarga-excel/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job || !job.buffer) return res.status(404).json({ error: 'Archivo no disponible.' });
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(job.buffer);
});

// Generador asíncrono de Excel
async function generarExcelAsync(params, jobId) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false, email } = params;
    const maxFilas = 3500;
    jobs[jobId].fase = "Preparando grupo y artículos";
    const workbookGrupos = XLSX.readFile('./grupos.xlsx');
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);

    const codigosGrupo = grupos.filter(row => row.grupo === grupo)
      .map(row => (row.codigo ? row.codigo.toString().trim() : null))
      .filter(Boolean);

    if (!codigosGrupo.length) {
      jobs[jobId].error = "No hay artículos para ese grupo.";
      jobs[jobId].progress = 100;
      return;
    }
    jobs[jobId].fase = "Descargando artículos base";
    const { usuario, password } = usuarios_api["Español"];
    const apiURL = "https://b2b.atosa.es:880/api/articulos/";
    let resp0;
    try {
      resp0 = await axios.get(apiURL, {
        auth: { username: usuario, password: password },
        timeout: 70000,
        httpsAgent: new https.Agent({ rejectUnauthorized: false }),
      });
    } catch (err) {
      jobs[jobId].error = "Error autenticando usuario principal: " + (err.response?.status || "") + " " + (err.response?.data || "");
      jobs[jobId].progress = 100;
      return;
    }
    let articulos_base = resp0.data
      .filter(art =>
        codigosGrupo.includes(art.codigo?.toString().trim()) &&
        (!soloStock || parseInt(art.disponible || 0) > 0)
      ).slice(0, maxFilas);

    if (!articulos_base.length) {
      jobs[jobId].error = "No hay artículos que coincidan con el filtro.";
      jobs[jobId].progress = 100;
      return;
    }

    //================= HISTÓRICO: SOLO SE AÑADE ESTA PARTE ==================
    const codigosActuales = articulos_base.map(art => (art.codigo ? art.codigo.toString().trim() : null)).filter(Boolean);
    const historico = loadHistorico();
    const codigosAnteriores = historico.codigos || [];
    const { altas, bajas } = compararListas(codigosAnteriores, codigosActuales);
    backupHistorico();
    saveHistorico(codigosActuales);
    //=======================================================================

    // === Resto de lógica de generación Excel ===
    jobs[jobId].fase = "Ordenando artículos según catálogo";
    articulos_base = ordenarArticulos(articulos_base);

    jobs[jobId].fase = "Descargando descripciones del idioma";
    let descripcionesIdioma = {};
    if (idioma !== "Español") {
      try {
        const userIdioma = usuarios_api[idioma];
        const respIdioma = await axios.get(apiURL, {
          auth: { username: userIdioma.usuario, password: userIdioma.password },
          timeout: 70000,
          httpsAgent: new https.Agent({ rejectUnauthorized: false }),
        });
        for (const art of respIdioma.data) {
          if (art.codigo && art.descripcion) {
            descripcionesIdioma[art.codigo.toString().trim()] = art.descripcion;
          }
        }
      } catch (e) {
        descripcionesIdioma = {};
      }
    }

    jobs[jobId].fase = "Calculando productos promocionales";
    let articulos_promocion = new Set();
    if (descuento > 0) {
      let precios0 = {}, precios8 = {};
      try {
        for (const art of articulos_base) {
          const cod = art.codigo ? art.codigo.toString().trim() : null;
          if (cod) precios0[cod] = parseFloat(art.precioVenta);
        }
        const resp8 = await axios.get(apiURL, {
          auth: { username: usuario8.usuario, password: usuario8.password },
          timeout: 70000,
          httpsAgent: new https.Agent({ rejectUnauthorized: false }),
        });
        for (const art of resp8.data) {
          const cod = art.codigo ? art.codigo.toString().trim() : null;
          if (cod) precios8[cod] = parseFloat(art.precioVenta);
        }
        for (const cod of Object.keys(precios0)) {
          if (precios8[cod] !== undefined && Math.abs(precios0[cod] - precios8[cod]) < 0.01) {
            articulos_promocion.add(cod);
          }
        }
      } catch {
        articulos_promocion = new Set();
      }
    }

    jobs[jobId].fase = "Componiendo Excel";
    const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
    const traducido = campos.map(c => diccionario_traduccion[idioma][c]);
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Listado');
    ws.addRow(traducido);
    const colWidths = { codigo: 11, descripcion: 30, disponible: 10, ean13: 10, precioVenta: 10, umv: 8, imagen: 15 };
    ws.columns = campos.map(c => ({ width: colWidths[c] || 15 }));

    const headerRow = ws.getRow(1);
    const cabeceraColor = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C3AED' } };
    headerRow.font = { bold: true, size: 15, color: { argb: 'FFFFFFFF' }, name: 'Segoe UI' };
    headerRow.height = filaAltura;
    campos.forEach((campo, idx) => {
      const cell = headerRow.getCell(idx + 1);
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true, textRotation: 0 };
      cell.fill = cabeceraColor;
      cell.border = { bottom: { style: 'thick', color: { argb: 'FF1E1E1E' } } };
    });

    const idxEAN = campos.indexOf("ean13") + 1;

    let pasoTotal = sinImagenes ? articulos_base.length : articulos_base.length * 2;
    let pasos = 0;

    for (const art of articulos_base) {
      const fila = [];
      const cod = art.codigo?.toString().trim();
      for (const campo of campos) {
        let valor = art[campo] ?? "";
        if (campo === "precioVenta") {
          if (descuento > 0 && !articulos_promocion.has(cod)) {
            valor = Math.round((parseFloat(valor) * (1 - descuento / 100)) * 100) / 100;
          } else {
            valor = parseFloat(valor);
          }
        } else if (campo === "descripcion" && idioma !== "Español") {
          if (descripcionesIdioma[cod]) valor = descripcionesIdioma[cod];
        }
        fila.push(valor);
      }
      ws.addRow(fila);
      pasos++;
      jobs[jobId].progress = Math.round((pasos / pasoTotal) * 97);
    }

    // Zebra y formato fila datos, EAN font 10 solo en datos
    for (let i = 2; i <= ws.rowCount; i++) {
      const row = ws.getRow(i);
      row.height = filaAltura;
      const zebra = i % 2 === 0 ? 'FFF3F4F6' : 'FFFFFFFF';
      for (let j = 1; j <= campos.length; j++) {
        const cell = row.getCell(j);
        const isEAN = j === idxEAN;
        const fontSize = isEAN ? 10 : 13;
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
          wrapText: campos[j - 1] === "descripcion",
          textRotation: isEAN ? 90 : 0
        };
        cell.font = { size: fontSize, name: 'Segoe UI' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: zebra } };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } }
        };
      }
    }

    if (!sinImagenes) {
      jobs[jobId].fase = "Insertando imágenes...";
      const limit = pLimit(3); // Control concurrencia
      const imagenPorDefecto = await crearImagenPorDefecto();
      const imagenesInsertadas = new Set();
      let imagenesExitosas = 0, imagenesConError = 0, imagenesDefault = 0;
      await Promise.all(articulos_base.map((art, i) => limit(async () => {
        let fotoBuffer = null;
        try {
          fotoBuffer = await obtenerFotoArticuloAPI(art.codigo, usuarios_api["Español"].usuario, usuarios_api["Español"].password, 3);
          if (!fotoBuffer || !validarBuffer(fotoBuffer)) {
            fotoBuffer = imagenPorDefecto; imagenesDefault++;
          } else {
            imagenesExitosas++;
          }
          const img = await Jimp.read(fotoBuffer);
          img.cover(imagenPx, imagenPx);
          const buffer = await img.getBufferAsync(Jimp.MIME_JPEG);
          const imgId = workbook.addImage({ buffer, extension: 'jpeg' });
          ws.addImage(imgId, {
            tl: { col: campos.length - 1, row: i + 1 },
            ext: { width: imagenPx, height: imagenPx }
          });
          imagenesInsertadas.add(i);
        } catch (error) {
          imagenesConError++;
        }
        pasos++;
        jobs[jobId].progress = Math.max(jobs[jobId].progress, Math.round((pasos / pasoTotal) * 99));
      })));
    }

    jobs[jobId].fase = "Finalizando";
    const buffer = await workbook.xlsx.writeBuffer();
    jobs[jobId].buffer = Buffer.from(buffer);
    jobs[jobId].progress = 100;
    jobs[jobId].filename = `listado_${grupo}_${idioma}${sinImagenes ? '_sinImagenes' : ''}.xlsx`;
    jobs[jobId].fase = "Completado";
    if (email) {
      jobs[jobId].fase = "Enviando email...";
      await enviarEmailConAdjunto(email, jobs[jobId].buffer, jobs[jobId].filename);
      jobs[jobId].fase = "Email enviado";
    }
  } catch (err) {
    jobs[jobId].error = `Error generando Excel: ${err.message}`;
    console.error('Error completo:', err);
    jobs[jobId].progress = 100;
    jobs[jobId].fase = "Error";
  }
}

// ============ ENDPOINT HISTÓRICO =============
app.get('/api/cambios-articulos', (req, res) => {
  const historico = loadHistorico();
  const codigosActuales = historico.codigos || [];
  let codigosAnteriores = [];
  if (fs.existsSync('./historico_anterior.json')) {
    codigosAnteriores = JSON.parse(fs.readFileSync('./historico_anterior.json', 'utf-8')).codigos || [];
  } else {
    codigosAnteriores = codigosActuales;
  }
  const { altas, bajas } = compararListas(codigosAnteriores, codigosActuales);
  res.json({
    resumen: {
      numAltas: altas.length,
      numBajas: bajas.length
    },
    nuevas: altas,
    bajas: bajas
  });
});

app.get('/', (req, res) => res.send('Servidor ATOSA backend funcionando.'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Escuchando en puerto ${PORT}`));

