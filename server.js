// server.js — ATOSA Excel con MongoDB Atlas y Panel Admin Cross-Domain

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

// ---------- CONFIGURACIÓN SEGURA CROSS-DOMAIN (Netlify + Render) ---------
app.set('trust proxy', 1);

app.use(cors({
  origin: 'https://webb2b.netlify.app',
  credentials: true
}));
app.use(express.json());

app.use(session({
  secret: process.env.ADMIN_SECRET || 'tu-clave-secreta-admin-2025',
  resave: false,
  saveUninitialized: false,
  proxy: true,
  cookie: {
    secure: true,       // TRUE: sólo si ambos son https (Netlify y Render)
    sameSite: 'none',   // ESENCIAL para cross-domain
    httpOnly: true,
    maxAge: 1000 * 60 * 60 * 2
  }
}));

// -----------------------------------------------------------------------

mongoose.connect(process.env.MONGODB_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
}).then(() => {
  console.log('Conectado a MongoDB Atlas exitosamente');
  cargarOrdenArticulos();
}).catch((error) => {
  console.error('Error conectando a MongoDB:', error);
  console.log('Intentando cargar desde archivos Excel como fallback...');
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
  Español: { codigo: "Código", descripcion: "Descripción", disponible: "Stock", ean13: "EAN", precioVenta: "Precio", umv: "UMV", imagen: "Imagen" },
  Inglés: { codigo: "Code", descripcion: "Description", disponible: "Available", ean13: "EAN", precioVenta: "Price", umv: "MOQ", imagen: "Image" },
  Francés: { codigo: "Code", descripcion: "Description", disponible: "Stock", ean13: "EAN", precioVenta: "Prix", umv: "MOQ", imagen: "Image" },
  Italiano: { codigo: "Codice", descripcion: "Descrizione", disponibile: "Stock", ean13: "EAN", precioVenta: "Prezzo", umv: "MOQ", imagen: "Immagine" }
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

// ========== FUNCIONES =====================
// (Pon aquí TODAS tus funciones largas tal como ya las tienes, sin cambiar nada)
// - crearBackup
// - cargarOrdenArticulos
// - cargarOrdenArticulosDesdeArchivo
// - ordenarArticulos
// - migrarDatosExistentes
// - obtenerFotoArticuloAPI
// - validarBuffer
// - crearImagenPorDefecto
// - enviarEmailConAdjunto
// - generarExcelAsync
// ...

// ========== ENDPOINTS ADMIN ===============
app.post('/admin/login', (req, res) => {
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
app.post('/admin/upload-grupos', requireAdmin, upload.single('archivo'), async (req, res) => {/* ...igual... */});
app.post('/admin/upload-orden', requireAdmin, upload.single('archivo'), async (req, res) => {/* ...igual... */});

// ========== ENDPOINTS PRINCIPALES ===========
app.get('/api/grupos', async (req, res) => {/* ...igual... */});
app.post('/api/genera-excel-final-async', async (req, res) => {/* ...igual... */});
app.get('/api/progreso/:jobId', (req, res) => {/* ...igual... */});
app.get('/api/descarga-excel/:jobId', (req, res) => {/* ...igual... */});

app.get('/', (req, res) => res.send('Servidor ATOSA backend funcionando.'));
app.listen(process.env.PORT || 3000, () => console.log(`Escuchando en puerto ${process.env.PORT || 3000}`));
