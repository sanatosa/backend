// server.js — ATOSA Excel con MongoDB Atlas y sistema admin seguro cross-domain

const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const https = require('https');
const { v4: uuidv4 } = require('uuid');
const Jimp = require('jimp');
const pLimit = require('p-limit').default;
const multer = require('multer');
const session = require('express-session');
const fs = require('fs');
const path = require('path');
const mongoose = require('mongoose');
require('dotenv').config();
const nodemailer = require('nodemailer');
const { Orden, Grupo } = require('./schemas');

const app = express();

// ----------- ¡¡ESTA ES LA CONFIGURACIÓN CRUCIAL!! -------------
app.set('trust proxy', 1);
// CORS SÓLO permite desde tu frontend y con credenciales
app.use(cors({
  origin: 'https://webb2b.netlify.app',
  credentials: true
}));
app.use(express.json());
// Session: proxy:true, secure:true en cookie, sameSite:'none', httpOnly:true
app.use(session({
  secret: process.env.ADMIN_SECRET || 'tu-clave-secreta-admin-2025',
  resave: false,
  saveUninitialized: false,
  proxy: true,
  cookie: {
    secure: true,    // SOLO TRUE si ambos, Netlify y Render, son HTTPS (que lo son)
    sameSite: 'none',
    httpOnly: true,
    maxAge: 1000 * 60 * 60 * 2
  }
}));
// ---------------------------------------------------------------

mongoose.connect(process.env.MONGODB_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => {
  console.log('Conectado a MongoDB Atlas exitosamente');
  cargarOrdenArticulos();
})
.catch((error) => {
  console.error('Error conectando a MongoDB:', error);
  cargarOrdenArticulosDesdeArchivo();
});

const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') cb(null, true);
    else cb(new Error('Solo se permiten archivos Excel (.xlsx)'), false);
  },
  limits: { fileSize: 5 * 1024 * 1024 }
});

const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'AtosaAdmin2025!';
function requireAdmin(req, res, next) {
  if (req.session.isAdmin) next();
  else res.status(401).json({ error: 'Acceso denegado. Inicia sesión como administrador.' });
}

const imagenPx = 110;
const filaAltura = 82.0;

const diccionario_traduccion = {
  Español: { codigo: "Código", descripcion: "Descripción", disponible: "Disponible", ean13: "EAN", precioVenta: "Precio", umv: "UMV", imagen: "Imagen" },
  Inglés: { codigo: "Code", descripcion: "Description", disponible: "Available", ean13: "EAN", precioVenta: "Price", umv: "MOQ", imagen: "Image" },
  Francés: { codigo: "Code", descripcion: "Description", disponible: "Disponible", ean13: "EAN", precioVenta: "Prix", umv: "MOQ", imagen: "Image" },
  Italiano: { codigo: "Codice", descripcion: "Descrizione", disponible: "Disponibile", ean13: "EAN", precioVenta: "Prezzo", umv: "MOQ", imagen: "Immagine" }
};

const usuarios_api = {
  Español: { usuario: "amazon@espana.es", password: "0glLD6g7Dg" },
  Inglés: { usuario: "ingles@atosa.es", password: "AtosaIngles" },
  Francés: { usuario: "frances@atosa.es", password: "AtosaFrances" },
  Italiano: { usuario: "italiano@atosa.es", password: "AtosaItaliano" }
};
const usuario8 = { usuario: "santi@tradeinn.com", password: "C8Zg1wqgfe" };
const jobs = {};
let ordenArticulos = {};

// ... TODAS TUS FUNCIONES IGUAL ...
// crearBackup, cargarOrdenArticulos, cargarOrdenArticulosDesdeArchivo, ordenarArticulos, migrarDatosExistentes, obtenerFotoArticuloAPI, validarBuffer, crearImagenPorDefecto, enviarEmailConAdjunto, generarExcelAsync

function crearBackup(archivo) {
  try {
    const fecha = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const backupDir = './backups';
    if (!fs.existsSync(backupDir)) fs.mkdirSync(backupDir);
    const extension = path.extname(archivo);
    const nombre = path.basename(archivo, extension);
    const backupPath = path.join(backupDir, `${nombre}_${fecha}${extension}`);
    fs.copyFileSync(archivo, backupPath);
    console.log(`Backup creado: ${backupPath}`);
    return backupPath;
  } catch (error) {
    console.error('Error creando backup:', error);
    return null;
  }
}

async function cargarOrdenArticulos() {
  try {
    const ordenData = await Orden.find().sort({ orden: 1 });
    ordenArticulos = {};
    ordenData.forEach(item => {
      ordenArticulos[item.codigo] = item.orden;
    });
    console.log(`Cargados ${Object.keys(ordenArticulos).length} artículos del orden desde MongoDB`);
  } catch (error) {
    console.error('Error cargando orden desde MongoDB:', error);
    cargarOrdenArticulosDesdeArchivo();
  }
}

function cargarOrdenArticulosDesdeArchivo() {
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
    console.log(`Cargados ${Object.keys(ordenArticulos).length} artículos del archivo orden.xlsx (fallback)`);
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

async function migrarDatosExistentes() {
  try {
    console.log('Iniciando migración de datos...');
    if (fs.existsSync('./orden.xlsx')) {
      const workbookOrden = XLSX.readFile('./orden.xlsx');
      const sheetOrden = workbookOrden.Sheets[workbookOrden.SheetNames[0]];
      const datosOrden = XLSX.utils.sheet_to_json(sheetOrden, { header: ['orden', 'codigo'] });
      await Orden.deleteMany({});
      const ordenItems = datosOrden.map(row => ({
        orden: parseInt(row.orden) || 999999,
        codigo: row.codigo ? row.codigo.toString().trim() : ''
      })).filter(item => item.codigo);
      await Orden.insertMany(ordenItems);
      console.log(`Migrados ${ordenItems.length} elementos de orden`);
    }
    if (fs.existsSync('./grupos.xlsx')) {
      const workbookGrupos = XLSX.readFile('./grupos.xlsx');
      const sheetGrupos = workbookGrupos.Sheets[workbookGrupos.SheetNames[0]];
      const datosGrupos = XLSX.utils.sheet_to_json(sheetGrupos);
      await Grupo.deleteMany({});
      const grupoItems = datosGrupos.map(row => ({
        grupo: row.grupo ? row.grupo.toString().trim() : '',
        codigo: row.codigo ? row.codigo.toString().trim() : ''
      })).filter(item => item.grupo && item.codigo);
      await Grupo.insertMany(grupoItems);
      console.log(`Migrados ${grupoItems.length} elementos de grupos`);
    }
    console.log('Migración completada exitosamente');
  } catch (error) {
    console.error('Error en la migración:', error);
  }
}

// ...el resto de funciones de imágenes, email, excel...

// === ENDPOINTS ADMIN ===
app.post('/admin/login', (req, res) => {
  // Diagnóstico: ¿la request llega como secure? (debes ver true en logs)
  console.log('req.secure:', req.secure);
  const { password } = req.body;
  if (password === ADMIN_PASSWORD) {
    req.session.isAdmin = true;
    res.json({ success: true, message: 'Acceso autorizado' });
  } else {
    res.status(401).json({ success: false, message: 'Contraseña incorrecta' });
  }
});

app.post('/admin/logout', (req, res) => {
  req.session.destroy(err => {
    if (err) res.status(500).json({ error: 'Error al cerrar sesión' });
    else res.json({ success: true, message: 'Sesión cerrada' });
  });
});

app.get('/admin/status', (req, res) => {
  res.json({ isAdmin: !!req.session.isAdmin });
});

app.post('/admin/upload-grupos', requireAdmin, upload.single('archivo'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No se subió ningún archivo' });
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const datos = XLSX.utils.sheet_to_json(sheet);
    if (datos.length === 0) return res.status(400).json({ error: 'El archivo está vacío' });
    const primeraFila = datos[0];
    if (!primeraFila.hasOwnProperty('grupo') || !primeraFila.hasOwnProperty('codigo')) {
      return res.status(400).json({ error: 'El archivo debe tener columnas "grupo" y "codigo"' });
    }
    if (fs.existsSync('./grupos.xlsx')) crearBackup('./grupos.xlsx');
    await Grupo.deleteMany({});
    const grupoItems = datos.map(row => ({
      grupo: row.grupo ? row.grupo.toString().trim() : '',
      codigo: row.codigo ? row.codigo.toString().trim() : ''
    })).filter(item => item.grupo && item.codigo);
    await Grupo.insertMany(grupoItems);
    fs.writeFileSync('./grupos.xlsx', req.file.buffer);
    console.log('Archivo grupos.xlsx actualizado por admin');
    res.json({
      success: true,
      message: 'Archivo grupos.xlsx actualizado correctamente en MongoDB',
      registros: grupoItems.length
    });
  } catch (error) {
    console.error('Error subiendo grupos.xlsx:', error);
    res.status(500).json({ error: 'Error procesando el archivo: ' + error.message });
  }
});

app.post('/admin/upload-orden', requireAdmin, upload.single('archivo'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No se subió ningún archivo' });
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const datos = XLSX.utils.sheet_to_json(sheet, { header: ['orden', 'codigo'] });
    if (datos.length === 0) return res.status(400).json({ error: 'El archivo está vacío' });
    const validRows = datos.filter(row => row.orden && row.codigo);
    if (validRows.length === 0) return res.status(400).json({ error: 'El archivo debe tener columnas de orden y código' });
    if (fs.existsSync('./orden.xlsx')) crearBackup('./orden.xlsx');
    await Orden.deleteMany({});
    const ordenItems = validRows.map(row => ({
      orden: parseInt(row.orden) || 999999,
      codigo: row.codigo ? row.codigo.toString().trim() : ''
    })).filter(item => item.codigo);
    await Orden.insertMany(ordenItems);
    fs.writeFileSync('./orden.xlsx', req.file.buffer);
    await cargarOrdenArticulos();
    console.log('Archivo orden.xlsx actualizado por admin');
    res.json({
      success: true,
      message: 'Archivo orden.xlsx actualizado correctamente en MongoDB',
      registros: ordenItems.length
    });
  } catch (error) {
    console.error('Error subiendo orden.xlsx:', error);
    res.status(500).json({ error: 'Error procesando el archivo: ' + error.message });
  }
});

// === ENDPOINTS PRINCIPALES ===
// ...todos los endpoints Excel frontend igual que ya tienes...

app.get('/api/grupos', async (req, res) => {
  try {
    const grupos = await Grupo.find();
    if (grupos.length > 0) {
      const nombres = [...new Set(grupos.map(g => g.grupo))].sort();
      res.json({ grupos: nombres });
    } else {
      const workbook = XLSX.readFile('./grupos.xlsx');
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const gruposData = XLSX.utils.sheet_to_json(sheet);
      const nombres = [...new Set(gruposData.map(row => (row.grupo ? row.grupo.toString().trim() : null)).filter(gr => gr && gr.length > 0))].sort();
      res.json({ grupos: nombres });
    }
  } catch (err) {
    console.error('Error obteniendo grupos:', err);
    res.status(500).json({ error: "No se pudieron obtener los grupos." });
  }
});

// ...el resto de endpoints siguientes igual...

app.get('/', (req, res) => res.send('Servidor ATOSA backend funcionando.'));
app.listen(process.env.PORT || 3000, () => console.log(`Escuchando en puerto ${process.env.PORT || 3000}`));
