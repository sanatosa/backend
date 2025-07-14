const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const https = require('https');

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
    res.status(500).send("No se pudo descargar el archivo de grupos.");
  }
});

// Descarga imagen desde URL y la devuelve como Buffer
async function descargarImagen(url) {
  if (!url) return null;
  try {
    const response = await axios.get(url, { responseType: 'arraybuffer', timeout: 15000 });
    return Buffer.from(response.data, 'binary');
  } catch (e) {
    console.log('No se pudo descargar imagen:', url);
    return null;
  }
}

// --- BLOQUE PRINCIPAL: GESTOR DE EXCEL ---
async function generaExcel(req, res) {
  try {
    const { grupo, idioma = "Español", descuento = 0, soloStock = false, maxFilas = 400 } = req.body;

    // 1. Leer grupos.xlsx desde GitHub
    const responseGrupos = await axios.get(
      'https://raw.githubusercontent.com/sanatosa/proxy/main/grupos.xlsx',
      { responseType: 'arraybuffer' }
    );
    const workbookGrupos = XLSX.read(responseGrupos.data, { type: 'buffer' });
    const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheetGrupos);

    // 2. Saca los códigos del grupo seleccionado
    const codigosGrupo = grupos.filter(row => row.grupo === grupo).map(row => row.codigo?.toString());
    if (!codigosGrupo.length) {
      return res.status(404).json({ error: "No hay artículos para ese grupo." });
    }

    // 3. Llama a la API de Atosa con usuario y password según idioma
    const { usuario, password } = usuarios_api[idioma] || usuarios_api["Español"];
    const apiURL = "https://b2b.atosa.es:880/api/articulos/";

    const respArticulos = await axios.get(apiURL, {
      auth: { username: usuario, password: password },
      timeout: 60_000,
      httpsAgent: new https.Agent({ rejectUnauthorized: false }),
    });

    let articulos = respArticulos.data;

    // 4. Filtra por grupo y stock (si procede)
    articulos = articulos.filter(art =>
      codigosGrupo.includes(art.codigo?.toString()) &&
      (!soloStock || parseInt(art.disponible || 0) > 0)
    );

    // 5. Limita resultados
    articulos = articulos.slice(0, maxFilas);

    if (!articulos.length) {
      return res.status(404).json({ error: "No hay artículos que coincidan con el filtro." });
    }

    // 6. Calcula artículos sin descuento (como en tu script)
    let articulos_sin_descuento = new Set();
    if (descuento > 0) {
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

    // 8. Añadir filas y descargar imágenes
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

      // Añadir imagen en la celda correspondiente (si hay URL en 'imagen')
      const urlImagen = art.imagen;
      if (urlImagen) {
        try {
          const imgBuffer = await descargarImagen(urlImagen);
          if (imgBuffer) {
            const ext = urlImagen.toLowerCase().endsWith('.png') ? 'png' : 'jpeg';
            const imageId = workbook.addImage({
              buffer: imgBuffer,
              extension: ext
            });
            // Columna imagen (última): campos.length - 1
            ws.addImage(imageId, {
              tl: { col: campos.length - 1, row: i + 1 },
              ext: { width: 110, height: 110 }
            });
          }
        } catch (e) {
          // Si la imagen no se puede añadir, no pasa nada
        }
      }
    }

    // 9. Formato de filas y celdas
    ws.eachRow({ includeEmpty: false }, function(row, rowNumber) {
      row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      row.height = 90;
      if (rowNumber === 1) row.font = { bold: true, size: 14 };
      else row.font = { size: 13 };
    });

    // 10. Devuelve el archivo Excel
    res.setHeader('Content-Disposition', `attachment; filename="listado_${grupo}_${idioma}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Error generando el Excel." });
  }
}

// Endpoint principal
app.post('/api/genera-excel-final', generaExcel);

// Alias para compatibilidad frontend antiguo
app.post('/api/genera-excel', generaExcel);

// Endpoint de prueba para saber si el backend está OK
app.get('/', (req, res) => {
  res.send('Servidor ATOSA backend funcionando.');
});

// Inicia el server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
});