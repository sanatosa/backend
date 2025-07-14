const express = require('express');
const axios = require('axios');
const cors = require('cors');
const XLSX = require('xlsx');
const bodyParser = require('body-parser');

const app = express();
app.use(cors());
app.use(bodyParser.json());

// --- ENDPOINT PROXY PARA GRUPOS.XLSX ---
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

// --- ENDPOINT PARA GENERAR EXCEL (EJEMPLO BÁSICO) ---
app.post('/api/genera-excel', async (req, res) => {
  try {
    const { grupo, idioma } = req.body;
    // Descarga el archivo de grupos desde GitHub
    const response = await axios.get(
      'https://raw.githubusercontent.com/sanatosa/proxy/main/grupos.xlsx',
      { responseType: 'arraybuffer' }
    );
    const workbook = XLSX.read(response.data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const grupos = XLSX.utils.sheet_to_json(sheet);

    // Filtra los artículos por grupo
    const articulosFiltrados = grupos.filter(row => row.grupo === grupo);

    if (!articulosFiltrados.length) {
      return res.status(404).json({ error: 'No hay artículos para ese grupo.' });
    }

    // Puedes aquí adaptar las columnas según idioma, etc.
    const nuevoLibro = XLSX.utils.book_new();
    const nuevaHoja = XLSX.utils.json_to_sheet(articulosFiltrados);
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'Artículos');
    const buffer = XLSX.write(nuevoLibro, { type: 'buffer', bookType: 'xlsx' });

    res.set('Content-Disposition', `attachment; filename=listado_${grupo}.xlsx`);
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Error generando el Excel.' });
  }
});

// --- ROOT ---
app.get('/', (req, res) => {
  res.send('Servidor ATOSA backend funcionando.');
});

// --- INICIA EL SERVER ---
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
});