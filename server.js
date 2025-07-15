// server.js — Excel profesional para tablet, descuentos robustos, promociones y traducción solo en descripción

const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const https = require('https');
const { v4: uuidv4 } = require('uuid');
const Jimp = require('jimp');
const pLimit = require('p-limit').default;

const app = express();
app.use(cors({ origin: 'https://webb2b.netlify.app' }));
app.use(express.json());

const diccionario_traduccion = {
  Español:   { codigo: "Código", descripcion: "Descripción", disponible: "Disponible", ean13: "EAN13", precioVenta: "Precio", umv: "UMV", imagen: "Imagen" },
  Inglés:    { codigo: "Code", descripcion: "Description", disponible: "Available", ean13: "EAN13", precioVenta: "Price", umv: "MOQ", imagen: "Image" },
  Francés:   { codigo: "Code", descripcion: "Description", disponible: "Disponible", ean13: "EAN13", precioVenta: "Prix", umv: "MOQ", imagen: "Image" },
  Italiano:  { codigo: "Codice", descripcion: "Descrizione", disponible: "Disponibile", ean13: "EAN13", precioVenta: "Prezzo", umv: "MOQ", imagen: "Immagine" }
};
const usuarios_api = {
  Español:   { usuario: "amazon@espana.es", password: "0glLD6g7Dg" },
  Inglés:    { usuario: "ingles@atosa.es", password: "AtosaIngles" },
  Francés:   { usuario: "frances@atosa.es", password: "AtosaFrances" },
  Italiano:  { usuario: "italiano@atosa.es", password: "AtosaItaliano" }
};
const usuario4 = { usuario: "compras@b2cmarketonline.es", password: "rXCRzzWKI6" };
const usuario8 = { usuario: "santi@tradeinn.com", password: "C8Zg1wqgfe" };
const jobs = {};

// --- ENDPOINTS ---
app.get('/api/grupos', async (req, res) => {
  try {
    const workbook = XLSX.readFile('./grupos.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheet);
    const nombres = [...new Set(
      grupos.map(row => (row.grupo ? row.grupo.toString().trim() : null)).filter(gr => gr && gr.length > 0)
    )].sort();
    res.json({ grupos: nombres });
  } catch (err) {
    res.status(500).json({ error: "No se pudieron obtener los grupos." });
  }
});

app.post('/api/genera-excel-final-async', async (req, res) => {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false } = req.body;
    const jobId = uuidv4();
    jobs[jobId] = { progress: 0, buffer: null, error: null, filename: null, startedAt: Date.now() };
    generarExcelAsync({ grupo, idioma, descuento, soloStock, sinImagenes }, jobId);
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
  res.json({ progress: job.progress, error: job.error, filename: job.filename, eta });
});

app.get('/api/descarga-excel/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job || !job.buffer) return res.status(404).json({ error: 'Archivo no disponible.' });
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(job.buffer);
});

// --- LÓGICA PRINCIPAL MEJORADA ---
async function generarExcelAsync(params, jobId) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false } = params;
    const maxFilas = 3500;
    const workbookGrupos = XLSX.readFile('./grupos.xlsx');
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    const codigosGrupo = grupos
      .filter(row => row.grupo === grupo)
      .map(row => (row.codigo ? row.codigo.toString().trim() : null))
      .filter(Boolean);

    if (!codigosGrupo.length) {
      jobs[jobId].error = "No hay artículos para ese grupo.";
      jobs[jobId].progress = 100;
      return;
    }

    // Artículos, precios y promociones en español
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
      jobs[jobId].error = "Error autenticando usuario principal de Atosa (Español): " + (err.response?.status || "") + " " + (err.response?.data || "");
      jobs[jobId].progress = 100;
      return;
    }

    const articulos_base = resp0.data
      .filter(art =>
        codigosGrupo.includes(art.codigo?.toString().trim()) &&
        (!soloStock || parseInt(art.disponible || 0) > 0) )
      .slice(0, maxFilas);

    if (!articulos_base.length) {
      jobs[jobId].error = "No hay artículos que coincidan con el filtro.";
      jobs[jobId].progress = 100;
      return;
    }

    // Descripciones en idioma diferente (si aplica)
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

    // --- Lógica robusta de promociones y descuentos ---
    let articulos_promocion = new Set();
    if (descuento > 0) {
      let precios0 = {}, precios4 = {}, precios8 = {};
      try {
        for(const art of articulos_base) {
          let cod = art.codigo ? art.codigo.toString().trim() : null;
          if (cod) precios0[cod] = parseFloat(art.precioVenta);
        }
        const [resp4, resp8] = await Promise.all([
          axios.get(apiURL, {
            auth: { username: usuario4.usuario, password: usuario4.password },
            timeout: 70000,
            httpsAgent: new https.Agent({ rejectUnauthorized: false }),
          }),
          axios.get(apiURL, {
            auth: { username: usuario8.usuario, password: usuario8.password },
            timeout: 70000,
            httpsAgent: new https.Agent({ rejectUnauthorized: false }),
          }),
        ]);
        for (const art of resp4.data) {
          let cod = art.codigo ? art.codigo.toString().trim() : null;
          if (cod) precios4[cod] = parseFloat(art.precioVenta);
        }
        for (const art of resp8.data) {
          let cod = art.codigo ? art.codigo.toString().trim() : null;
          if (cod) precios8[cod] = parseFloat(art.precioVenta);
        }
        for (const cod of Object.keys(precios0)) {
          if (
            precios4[cod] !== undefined &&
            precios8[cod] !== undefined &&
            Math.abs(precios4[cod] - precios8[cod]) < 0.01
          ) {
            articulos_promocion.add(cod);
          }
        }
      } catch {
        articulos_promocion = new Set();
      }
    }

    // --- FORMATO EXCEL PRO ---
    const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
    const traducido = campos.map(c => diccionario_traduccion[idioma][c]);
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Listado');
    ws.addRow(traducido);

    const colWidths = { codigo: 12, descripcion: 40, disponible: 12, ean13: 12, precioVenta: 12, umv: 10, imagen: 18 };
    ws.columns = campos.map(c => ({ width: colWidths[c] || 15 }));

    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, size: 15, color: { argb: 'FFFFFFFF' }, name: 'Segoe UI' };
    headerRow.height = 90;
    headerRow.eachCell(cell => {
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true, textRotation: campos[cell.col - 1] === "ean13" ? 90 : 0 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1976D2' }};
      cell.border = {bottom: {style: 'thick', color: {argb:'FF1E1E1E'}}};
    });

    let pasoTotal = sinImagenes ? articulos_base.length : articulos_base.length * 2;
    let pasos = 0;
    let failedFotos = [];
    let zebraColors = ['FFFFFFFF','FFF3F4F6'];

    articulos_base.forEach((art, i) => {
      const fila = [];
      const cod = art.codigo?.toString().trim();
      campos.forEach(campo => {
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
      });
      ws.addRow(fila);
      pasos++;
      jobs[jobId].progress = Math.round((pasos / pasoTotal) * 98);
    });

    // Ajuste visual fila a fila: zebra, precios, stocks y promo
    for (let i = 2; i <= ws.rowCount; i++) {
      const row = ws.getRow(i);
      row.height = 90;
      row.font = { size: 13, name: 'Segoe UI' };
      const zebraColor = { type: 'pattern', pattern: 'solid', fgColor: { argb: zebraColors[(i%2)] } };
      let cod = row.getCell(1).value?.toString().trim();
      let esPromo = articulos_promocion.has(cod);
      let stock = Number(row.getCell(3).value || 0);
      for (let j = 1; j <= campos.length; j++) {
        let cell = row.getCell(j);
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true, textRotation: (campos[j-1] === "ean13" ? 90 : 0)};
        cell.fill = zebraColor;
        cell.border = {top:{style:'thin',color:{argb:'FFCCCCCC'}},bottom:{style:'thin',color:{argb:'FFCCCCCC'}}};
        // Colorea Precio de promociones
        if (campos[j-1] === "precioVenta" && esPromo) {
          cell.fill = { type: 'pattern', pattern:'solid', fgColor: {argb: 'FFBBDEFB'} };
          cell.font = {...cell.font, color: {argb: 'FF1565C0'}, italic: true };
          cell.value = cell.value != null ? `${cell.value} PROMO` : 'PROMO';
        }
        // Stock bajo
        if (campos[j-1] === "disponible") {
          if (stock <= 10) {
            cell.fill = { type: 'pattern', pattern:'solid', fgColor: {argb:'FFFFCDD2'} };
            cell.font = {...cell.font, color: {argb:'FFD32F2F'}, bold:true };
          } else if (stock <= 20) {
            cell.fill = { type: 'pattern', pattern:'solid', fgColor: {argb:'FFFFF9C4'} };
            cell.font = {...cell.font, color: {argb:'FFF9A825'}, bold:true };
          }
        }
        // Precio normal
        if (campos[j-1] === "precioVenta" && !esPromo) {
          cell.font = {...cell.font, color: {argb:'FF1976D2'}, bold:true};
          cell.alignment = {...cell.alignment, horizontal: 'right'};
        }
      }
    }

    // Imágenes igual que siempre
    if (!sinImagenes) {
      const limit = pLimit(3);
      await Promise.all(articulos_base.map((art, i) => limit(async () => {
        let fotoBuffer = await obtenerFotoArticuloAPI(art.codigo, usuario, password, 2);
        if (fotoBuffer) {
          try {
            let img = await Jimp.read(fotoBuffer);
            img.cover(110, 110);
            img.quality(60);
            const miniBuffer = await img.getBufferAsync(Jimp.MIME_JPEG);
            const imageId = workbook.addImage({ buffer: miniBuffer, extension: 'jpeg' });
            ws.addImage(imageId, {
              tl: { col: campos.length - 1, row: i + 1 },
              ext: { width: 110, height: 110 }
            });
          } catch (e) {
            failedFotos.push(art.codigo?.toString());
          }
        } else {
          failedFotos.push(art.codigo?.toString());
        }
        pasos++;
        if (jobs[jobId].progress < 99) {
          jobs[jobId].progress = Math.max(jobs[jobId].progress, Math.round((pasos / pasoTotal) * 99));
        }
      })));
    }

    let buffer = await workbook.xlsx.writeBuffer();
    jobs[jobId].buffer = Buffer.from(buffer);
    jobs[jobId].progress = 100;
    jobs[jobId].filename = `listado_${grupo}_${idioma}${sinImagenes ? '_sinImagenes' : ''}.xlsx`;
  } catch (err) {
    jobs[jobId].error = "Error generando el Excel (excepción interna).";
    jobs[jobId].progress = 100;
    console.error("[generarExcelAsync] Error:", err);
  }
}

async function obtenerFotoArticuloAPI(codigo, usuario, password, intentos = 2) {
  for (let i = 0; i < intentos; i++) {
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
    } catch (e) {
      await new Promise(res => setTimeout(res, 500));
    }
  }
  return null;
}

app.get('/', (req, res) => res.send('Servidor ATOSA backend funcionando.'));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor escuchando en puerto ${PORT}`));
