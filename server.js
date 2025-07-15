const express = require('express');
const cors = require('cors');
const axios = require('axios');
const https = require('https');
const XLSX = require('xlsx');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

// Configuración de la API de ATOSA
const API_BASE_URL = 'https://b2b.atosa.es:880/api';
const API_CREDENTIALS = { username: 'amazon@espana.es', password: '0glLD6g7Dg' };

// Endpoint: Obtener grupos desde grupos.xlsx
app.get('/api/grupos', (req, res) => {
  try {
    const workbook = XLSX.readFile('./grupos.xlsx');
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    res.json(data);
  } catch (error) {
    console.error('Error al leer grupos.xlsx:', error.message);
    res.status(500).json({ error: 'Error al leer grupos.xlsx.' });
  }
});

// Endpoint: Obtener artículos por grupo desde la API
app.get('/api/articulos/grupo/:codigo', async (req, res) => {
  const { codigo } = req.params;
  try {
    const response = await axios.get(`${API_BASE_URL}/articulos/grupo/${codigo}`, {
      auth: API_CREDENTIALS,
      httpsAgent: new https.Agent({ rejectUnauthorized: false })
    });
    res.json(response.data);
  } catch (error) {
    console.error(`Error al obtener artículos del grupo ${codigo}:`, error.message);
    res.status(500).json({ error: `Error al obtener artículos del grupo ${codigo}.` });
  }
});

// Endpoint: Obtener fotos de un artículo desde la API
app.get('/api/articulos/foto/:codigo', async (req, res) => {
  const { codigo } = req.params;
  try {
    const response = await axios.get(`${API_BASE_URL}/articulos/foto/${codigo}`, {
      auth: API_CREDENTIALS,
      httpsAgent: new https.Agent({ rejectUnauthorized: false })
    });
    const fotos = response.data.fotos || [];
    if (fotos.length === 0) {
      return res.status(404).json({ error: `No hay fotos disponibles para el artículo ${codigo}.` });
    }
    res.json(fotos[0]); // Devolver solo la primera foto en Base64
  } catch (error) {
    console.error(`Error al obtener fotos del artículo ${codigo}:`, error.message);
    res.status(500).json({ error: `Error al obtener fotos del artículo ${codigo}.` });
  }
});

// Inicializar el servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('Servidor ATOSA backend funcionando.');
});