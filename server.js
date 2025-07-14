const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const https = require('https');
const { v4: uuidv4 } = require('uuid');
const Jimp = require('jimp');

const app = express();
app.use(cors());
app.use(express.json());

const diccionario_traduccion = {
  Español: {
    codigo: "Código",
    descripcion: "Descripción",
    disponible: "Disponible",
    ean13: "EAN13",
    precioVenta: "Precio",
    umv: "UMV",
    imagen: "Imagen",
  },
  Inglés: {
    codigo: "Code",
    descripcion: "Description",
    disponible: "Available",
    ean13: "EAN13",
    precioVenta: "Price",
    umv: "MOQ",
    imagen: "Image",
  },
  Francés: {
    codigo: "Code",
    descripcion: "Description",
    disponible: "Disponible",
    ean13: "EAN13",
    precioVenta: "Prix",
    umv: "MOQ",
    imagen: "Image",
  },
  Italiano: {
    codigo: "Codice",
    descripcion: "Descrizione",
    disponible: "Disponibile",
    ean13: "EAN13",
    precioVenta: "Prezzo",
    umv: "MOQ",
    imagen: "Immagine",
  }
};

const usuarios_api = {
  Español: { usuario: "amazon@espana.es", password: "0glLD6g7Dg" },
  Inglés: { usuario: "ingles@atosa.es", password: "AtosaIngles" },
  Francés: { usuario: "frances@atosa.es", password: "AtosaFrances" },
  Italiano: { usuario: "italiano@atosa.es", password: "AtosaItaliano" }
};

// Endpoint para descargar grupos.xlsx del repositorio GitHub (proxy)
app.get('/grupos.xlsx', async (req, res) => {
  try {
    const response = await axios.get(
      'https://raw.githubusercontent.com/sanatosa/proxy/main/grupos.xlsx',
      { responseType: 'arraybuffer' }
    );
    res.set('Access-Control-Allow-Origin', '*');
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(response.data);
  } catch (err) {
    res.status(500).json({ error: "No se pudo descargar el archivo de grupos." });
  }
});

// Endpoint para obtener la lista de grupos únicos (para el desplegable)
app.get('/api/grupos', async (req, res) => {
  try {
    const response = await axios.get(
      'https://raw.githubusercontent.com/sanatosa/proxy/main/grupos.xlsx',
      { responseType: 'arraybuffer' }
    );
    const workbookGrupos = XLSX.read(response.data, { type: 'buffer' });
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    const nombres = [...new Set(grupos.map(row => row.grupo).filter(Boolean))].sort();
    res.json({ grupos: nombres });
  } catch (err) {
    res.status(500).json({ error: "No se pudieron obtener los grupos." });
  }
});

// Obtener la primera foto de un artículo en Buffer (jpeg)
async function obtenerFotoArticulo(codigo, usuario, password) {
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
    console.log(`No se pudo obtener foto para ${codigo}:`, e.message);
    return null;
  }
}

// ---- GESTIÓN DE JOBS Y PROGRESO EN MEMORIA ----
const jobs = {}; // jobId: { progress, buffer, error, filename }

app.post('/api/genera-excel-final-async', async (req, res) => {
  try {
    const { grupo } = req.body;
    if (!grupo || typeof grupo !== 'string' || grupo.trim() === '') {
      return res.status(400).json({ error: "El parámetro 'grupo' es obligatorio." });
    }

    const jobId = uuidv4();
    jobs[jobId] = { progress: 0, buffer: null, error: null, filename: null };

    console.log(`[${new Date().toISOString()}] Nueva petición de Excel para grupo "${grupo}", jobId: ${jobId}`);

    generarExcelAsync(req.body, jobId);

    res.json({ jobId });
  } catch (err) {
    res.status(500).json({ error: "Error iniciando la generación del Excel." });
  }
});

// Endpoint para consultar progreso
app.get('/api/progreso/:jobId', (req, res) => {
  const { jobId } = req.params;
  if (!jobs[jobId]) return res.status(404).json({ error: 'Trabajo no encontrado' });
  res.json({
    progress: jobs[jobId].progress,
    error: jobs[jobId].error,
    filename: jobs[jobId].filename
  });
});

// Endpoint para descargar el archivo generado
app.get('/api/descarga-excel/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job || !job.buffer) {
    return res.status(404).json({ error: 'Archivo no disponible.' });
  }
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(job.buffer);
});

// -------- GENERADOR EXCEL ASÍNCRONO -----------
async function generarExcelAsync(params, jobId) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, maxFilas = 400 } = params;

    // 1. Leer grupos.xlsx desde GitHub
    let responseGrupos;
    try {
      responseGrupos = await axios.get(
        'https://raw.githubusercontent.com/sanatosa/proxy/main/grupos.xlsx',
        { responseType: 'arraybuffer' }
      );
    } catch (err) {
      jobs[jobId].error = "No se pudo descargar el archivo de grupos.";
      jobs[jobId].progress = 100;
      return;
    }

    let workbookGrupos, sheetGrupos, grupos;
    try {
      workbookGrupos = XLSX.read(responseGrupos.data, { type: 'buffer' });
      sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
      grupos = XLSX.utils.sheet_to_json(sheetGrupos);
    } catch (err) {
      jobs[jobId].error = "No se pudo leer el archivo de grupos.";
      jobs[jobId].progress = 100;
      return;
    }

    // 2. Saca los códigos del grupo seleccionado
    const codigosGrupo = grupos.filter(row => row.grupo === grupo).map(row => row.codigo?.toString());
    if (!codigosGrupo.length) {
      jobs[jobId].error = "No hay artículos para ese grupo.";
      jobs[jobId].progress = 100;
      return;
    }

    // 3. Llama a la API de Atosa con usuario y password según idioma
    const { usuario, password } = usuarios_api[idioma] || usuarios_api["Español"];
    const apiURL = "https://b2b.atosa.es:880/api/articulos/";

    let respArticulos, articulos;
    try {
      respArticulos = await axios.get(apiURL, {
        auth: { username: usuario, password: password },
        timeout: 60_000,
        httpsAgent: new https.Agent({ rejectUnauthorized: false }),
      });
      articulos = respArticulos.data;
    } catch (err) {
      jobs[jobId].error = "No se pudo conectar con la API de Atosa.";
      jobs[jobId].progress = 100;
      return;
    }

    // 4. Filtra por grupo y stock (si procede)
    articulos = articulos.filter(art =>
      codigosGrupo.includes(art.codigo?.toString()) &&
      (!soloStock || parseInt(art.disponible || 0) > 0)
    );

    // 5. Limita resultados
    articulos = articulos.slice(0, maxFilas);

    if (!articulos.length) {
      jobs[jobId].error = "No hay artículos que coincidan con el filtro.";
      jobs[jobId].progress = 100;
      return;
    }

    // 6. Calcula artículos sin descuento (como en tu script)
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
            timeout: 60_000,
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
      } catch (err) {
        jobs[jobId].error = "No se pudo calcular los descuentos.";
        jobs[jobId].progress = 100;
        return;
      }
    }

    // 7. Crear Excel
    const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
    const traducido = campos.map(c => diccionario_traduccion[idioma][c]);

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Listado');

    ws.addRow(traducido);

    // Ancho y formato de columnas
    const colWidths = [12, 40, 12, 12, 12, 10, 18];
    ws.columns = ws.columns.map((col, idx) => ({
      ...col,
      width: colWidths[idx] || 15
    }));

    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // 8. Añadir filas y descargar imágenes de la API oficial
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

      // Obtener imagen desde la API de fotos oficial
      const fotoBuffer = await obtenerFotoArticulo(art.codigo, usuario, password);
if (fotoBuffer) {
  // REDIMENSIONAR Y COMPRIMIR como en Python
  const img = await Jimp.read(fotoBuffer);
  img.resize(110, 110).quality(60); // Igual que tu script Python
  const miniBuffer = await img.getBufferAsync(Jimp.MIME_JPEG);

  const imageId = workbook.addImage({
    buffer: miniBuffer,
    extension: 'jpeg'
  });
  ws.addImage(imageId, {
    tl: { col: campos.length - 1, row: i + 1 },
    ext: { width: 110, height: 110 }
  });
}
      jobs[jobId].progress = Math.round(((i + 1) / articulos.length) * 80);
    }

    // 9. Formato de filas y celdas
    ws.eachRow({ includeEmpty: false }, function(row, rowNumber) {
      row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      row.height = 90;
      if (rowNumber === 1) row.font = { bold: true, size: 14 };
      else row.font = { size: 13 };
    });

    // 10. Devuelve el archivo Excel (guarda en memoria)
    let buffer;
    try {
      buffer = await workbook.xlsx.writeBuffer();
    } catch (err) {
      jobs[jobId].error = "No se pudo generar el archivo Excel.";
      jobs[jobId].progress = 100;
      return;
    }
    jobs[jobId].buffer = Buffer.from(buffer);
    jobs[jobId].progress = 100;
    jobs[jobId].filename = `listado_${grupo}_${idioma}.xlsx`;
    console.log(`[${new Date().toISOString()}] Terminado jobId ${jobId} (${articulos.length} artículos)`);
  } catch (err) {
    console.error(err);
    jobs[jobId].error = "Error generando el Excel (excepción interna).";
    jobs[jobId].progress = 100;
  }
}

// Endpoint de prueba para saber si el backend está OK
app.get('/', (req, res) => {
  res.send('Servidor ATOSA backend funcionando.');
});

// Inicia el server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
});