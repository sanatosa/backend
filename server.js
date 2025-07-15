// server.js — Backend ATOSA con formato Excel idéntico al script Python y lógica exacta de descuentos/promociones

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

app.use(cors({ origin: 'https://webb2b.netlify.app' })); // Ajusta si tu frontend cambia
app.use(express.json());

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
const usuariosDescuento = [
  { usuario: "compras@b2cmarketonline.es", password: "rXCRzzWKI6" },
  { usuario: "santi@tradeinn.com", password: "C8Zg1wqgfe" }
];

const jobs = {};

// ENDPOINTS
app.get('/api/grupos', async (req, res) => {
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

// LÓGICA DE GENERACIÓN DE EXCEL: Formato igual que el script y descuentos sin errores
async function generarExcelAsync(params, jobId) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false } = params;
    const maxFilas = 3500;
    const workbookGrupos = XLSX.readFile('./grupos.xlsx');
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    const codigosGrupo = grupos
      .filter(row => row.grupo === grupo)
      .map(row => (row.codigo ? row.codigo.toString() : null))
      .filter(Boolean);

    if (!codigosGrupo.length) {
      jobs[jobId].error = "No hay artículos para ese grupo.";
      jobs[jobId].progress = 100;
      return;
    }

    const { usuario, password } = usuarios_api[idioma] || usuarios_api["Español"];
    const apiURL = "https://b2b.atosa.es:880/api/articulos/";

    const respArticulos = await axios.get(apiURL, {
      auth: { username: usuario, password: password },
      timeout: 70000,
      httpsAgent: new https.Agent({ rejectUnauthorized: false }),
    });

    const articulos = respArticulos.data
      .filter(art =>
        codigosGrupo.includes(art.codigo?.toString()) &&
        (!soloStock || parseInt(art.disponible || 0) > 0)
      )
      .slice(0, maxFilas);

    if (!articulos.length) {
      jobs[jobId].error = "No hay artículos que coincidan con el filtro.";
      jobs[jobId].progress = 100;
      return;
    }

    // Lógica exacta de descuentos/promociones según script Python
    let articulos_sin_descuento = new Set();
    if (descuento > 0) {
      try {
        const [resp4, resp8] = await Promise.all(usuariosDescuento.map(u =>
          axios.get(apiURL, {
            auth: { username: u.usuario, password: u.password },
            timeout: 70000,
            httpsAgent: new https.Agent({ rejectUnauthorized: false }),
          })
        ));

        const precios0 = Object.fromEntries(articulos.map(a => [a.codigo.toString(), parseFloat(a.precioVenta)]));
        const precios4 = Object.fromEntries(resp4.data.map(a => [a.codigo.toString(), parseFloat(a.precioVenta)]));
        const precios8 = Object.fromEntries(resp8.data.map(a => [a.codigo.toString(), parseFloat(a.precioVenta)]));

        for (const codigo of Object.keys(precios0)) {
          if (
            precios4[codigo] !== undefined &&
            precios8[codigo] !== undefined &&
            Math.abs(precios0[codigo] - precios4[codigo]) < 0.01 &&
            Math.abs(precios0[codigo] - precios8[codigo]) < 0.01
          ) {
            articulos_sin_descuento.add(codigo);
          }
        }
      } catch (err) {
        articulos_sin_descuento = new Set();
      }
    }

    // Formato idéntico a openpyxl / script Python
    const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
    const traducido = campos.map(c => diccionario_traduccion[idioma][c]);
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Listado');
    ws.addRow(traducido);

    // ANCHOS de columna idénticos
    const colWidths = { codigo: 12, descripcion: 40, disponible: 12, ean13: 12, precioVenta: 12, umv: 10, imagen: 18 };
    ws.columns = campos.map(c => ({ width: colWidths[c] || 15 }));

    // Cabecera: fuente, altura, alineación, wrapText y rotación en EAN13
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, size: 14 };
    headerRow.height = 90;

    campos.forEach((campo, idx) => {
      const cell = headerRow.getCell(idx + 1);
      if (campo === "descripcion") {
        cell.alignment = { wrapText: true, vertical: "middle", horizontal: "center" };
      } else if (campo === "ean13") {
        cell.alignment = { textRotation: 90, horizontal: "center", vertical: "center" };
      } else {
        cell.alignment = { horizontal: "center", vertical: "center" };
      }
    });

    let pasoTotal = sinImagenes ? articulos.length : articulos.length * 2;
    let pasos = 0;
    let failedFotos = [];

    // Agrega filas
    articulos.forEach((art, i) => {
      const fila = [];
      campos.forEach(campo => {
        let valor = art[campo] ?? "";
        if (campo === "precioVenta") {
          try {
            if (
              descuento > 0 &&
              !articulos_sin_descuento.has(art.codigo?.toString())
            ) {
              valor = Math.round((parseFloat(valor) * (1 - descuento / 100)) * 100) / 100;
            } else {
              valor = parseFloat(valor);
            }
          } catch {}
        }
        fila.push(valor);
      });
      ws.addRow(fila);
      pasos++;
      jobs[jobId].progress = Math.round((pasos / pasoTotal) * 98);
    });

    // Aplica formatos: altura, fuente y alineación a todas las filas, incluida la rotación EAN13
    for (let i = 2; i <= ws.rowCount; i++) {
      const row = ws.getRow(i);
      row.height = 90;
      row.font = { size: 13 };
      campos.forEach((campo, idx) => {
        const cell = row.getCell(idx + 1);
        if (campo === "descripcion") {
          cell.alignment = { wrapText: true, vertical: "middle", horizontal: "center" };
        } else if (campo === "ean13") {
          cell.alignment = { textRotation: 90, horizontal: "center", vertical: "center" };
        } else {
          cell.alignment = { horizontal: "center", vertical: "center" };
        }
      });
    }

    // Inserta imágenes si corresponde (idéntica lógica a versiones previas)
    if (!sinImagenes) {
      const limit = pLimit(3);
      await Promise.all(articulos.map((art, i) => limit(async () => {
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
