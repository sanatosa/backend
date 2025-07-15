const express = require('express');
const axios = require('axios');
const cors = require('cors');
const https = require('https');

const app = express();
app.use(cors());
app.use(express.json());

// Configuración de la API de ATOSA
const API_BASE_URL = 'https://b2b.atosa.es:880/api';
const API_CREDENTIALS = { username: 'amazon@espana.es', password: '0glLD6g7Dg' };

// Endpoint: Obtener grupos
app.get('/api/grupos', async (req, res) => {
  try {
    const response = await axios.get(`${API_BASE_URL}/grupos/`, {
      auth: API_CREDENTIALS,
      httpsAgent: new https.Agent({ rejectUnauthorized: false })
    });
    res.json(response.data);
  } catch (error) {
    console.error('Error al obtener grupos:', error.message);
    res.status(500).json({ error: 'Error al obtener grupos.' });
  }
});

// Endpoint: Obtener artículos por grupo
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

// Inicializar el servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('Servidor ATOSA backend funcionando.');
});