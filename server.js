// server.js — versión corregida: inserción precisa de imágenes, sincronización por fila real

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

const imagenPx = 110;
const filaAltura = 82.0; // Altura exacta en puntos para la fila Excel

const diccionario_traduccion = { /* ... igual que antes ... */ };
const usuarios_api = { /* ... igual que antes ... */ };
const usuario8 = { /* ... igual que antes ... */ };

const jobs = {};

app.post('/api/genera-excel-final-async', async (req, res) => {
    try {
        const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false } = req.body;
        const jobId = uuidv4();
        jobs[jobId] = { progress: 0, buffer: null, error: null, filename: null, startedAt: Date.now(), fase: "Preparando" };
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

// Clave ÚNICA imagen: por modelo y tipo (niño/adulto)
function claveFoto(descripcion) {
    let texto = (descripcion || '').trim().replace(/\s+/g, ' ').toUpperCase();
    let match = texto.match(/^(.+?)\s+((?:[0-9]+(?:-[0-9]+)?)|(?:XS|S|M|L|XL|XXL|XS-S|S-M|M-L|L-XL|XL-XXL))$/);
    if (!match) return texto;
    return match[1].trim() + '___' + (/^\d/.test(match[2]) ? 'NINO' : 'ADULTO');
}

async function obtenerFotoArticuloAPI(codigo, usuario, password, intentos = 2) {
    for (let i = 0; i < intentos; i++) {
        try {
            const resp = await axios.get(`https://b2b.atosa.es:880/api/articulos/foto/${codigo}`, {
                auth: { username: usuario, password },
                timeout: 10000,
                httpsAgent: new https.Agent({ rejectUnauthorized: false }),
            });
            const fotos = resp.data.fotos;
            if (Array.isArray(fotos) && fotos.length > 0) {
                return Buffer.from(fotos[0], 'base64');
            }
        } catch (e) {
            await new Promise(r => setTimeout(r, 500));
        }
    }
    return null;
}

async function generarExcelAsync(params, jobId) {
    try {
        const { grupo, idioma = "Español", descuento = 0, soloStock = false, sinImagenes = false } = params;
        const maxFilas = 3500;
        jobs[jobId].fase = "Preparando grupo y artículos";

        // Grupos y artículos base
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
        const articulos_base = resp0.data
            .filter(art =>
                codigosGrupo.includes(art.codigo?.toString().trim()) &&
                (!soloStock || parseInt(art.disponible || 0) > 0))
            .slice(0, maxFilas);
        if (!articulos_base.length) {
            jobs[jobId].error = "No hay artículos que coincidan con el filtro.";
            jobs[jobId].progress = 100;
            return;
        }

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
            } catch {
                descripcionesIdioma = {};
            }
        }

        // 2 PASADAS PARA IMÁGENES, AHORA CON MAPEO DE FILAS (sin desfases)
        jobs[jobId].fase = "Mapeando claves de imagen...";
        const clavesFila = articulos_base.map(art => claveFoto(art.descripcion));
        const clavesUnicas = Array.from(new Set(clavesFila));
        jobs[jobId].fase = "Descargando imágenes únicas...";
        const fotosPorModelo = {};
        const limitFoto = pLimit(5);
        await Promise.all(clavesUnicas.map(clave =>
            limitFoto(async () => {
                const idxArt = clavesFila.findIndex(c => c === clave);
                if (idxArt !== -1) {
                    const codigo = articulos_base[idxArt].codigo;
                    fotosPorModelo[clave] = await obtenerFotoArticuloAPI(
                        codigo,
                        usuario,
                        password,
                        2
                    );
                } else {
                    fotosPorModelo[clave] = null;
                }
            })
        ));

        jobs[jobId].fase = "Componiendo Excel";
        const campos = ["codigo", "descripcion", "disponible", "ean13", "precioVenta", "umv", "imagen"];
        const traducido = campos.map(c => diccionario_traduccion[idioma][c]);
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Listado');
        ws.addRow(traducido);
        const colWidths = { codigo: 11, descripcion: 30, disponible: 10, ean13: 10, precioVenta: 10, umv: 8, imagen: 15 };
        ws.columns = campos.map(c => ({ width: colWidths[c] || 15 }));

        // CABECERA
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
        let pasoTotal = articulos_base.length + (sinImagenes ? 0 : articulos_base.length), pasos = 0;

        // MAPEO entre código y fila real en Excel
        const filaPorCodigo = {};

        // Filas de datos y registro del número de fila
        for (const art of articulos_base) {
            const fila = [];
            const cod = art.codigo?.toString().trim();
            for (const campo of campos) {
                let valor = art[campo] ?? "";
                if (campo === "precioVenta") {
                    if (descuento > 0 && !articulos_promocion?.has?.(cod)) {
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
            filaPorCodigo[cod] = ws.lastRow.number;
            pasos++;
            jobs[jobId].progress = Math.round((pasos / pasoTotal) * 96);
        }

        // FORMATO filas y zebra
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

        // INSERCIÓN DE FOTOS sincronizada con fila real
        if (!sinImagenes) {
            jobs[jobId].fase = "Insertando imágenes...";
            const limit = pLimit(5);
            await Promise.all(articulos_base.map((art) => limit(async () => {
                const clave = claveFoto(art.descripcion);
                const fotoBuffer = fotosPorModelo[clave];
                const cod = art.codigo?.toString().trim();
                if (fotoBuffer && filaPorCodigo[cod]) {
                    try {
                        const img = await Jimp.read(fotoBuffer);
                        img.cover(imagenPx, imagenPx);
                        const buffer = await img.getBufferAsync(Jimp.MIME_JPEG);
                        const imgId = workbook.addImage({ buffer, extension: 'jpeg' });
                        ws.addImage(imgId, {
                            tl: { col: campos.length - 1, row: filaPorCodigo[cod] },
                            ext: { width: imagenPx, height: imagenPx }
                        });
                    } catch { /* error de imagen, saltar */ }
                }
                pasos++;
                jobs[jobId].progress = Math.max(jobs[jobId].progress, Math.round((pasos / pasoTotal) * 99));
            })));
        }

        jobs[jobId].fase = "Finalizando";
        const buffer = await workbook.xlsx.writeBuffer();
        jobs[jobId].buffer = Buffer.from(buffer);
        jobs[jobId].progress = 100;
        jobs[jobId].filename = `listado_${params.grupo}_${params.idioma}${sinImagenes ? '_sinImagenes' : ''}.xlsx`;
        jobs[jobId].fase = "Completado";
    } catch (err) {
        jobs[jobId].error = "Error generando el Excel.";
        jobs[jobId].progress = 100;
        jobs[jobId].fase = "Error";
        console.error(err);
    }
}

app.get('/', (req, res) => res.send('Servidor ATOSA backend funcionando.'));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Escuchando en puerto ${PORT}`));
