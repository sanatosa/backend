const mongoose = require('mongoose');

// Esquema para el orden de artículos
const ordenSchema = new mongoose.Schema({
  orden: {
    type: Number,
    required: true,
    index: true
  },
  codigo: {
    type: String,
    required: true,
    unique: true,
    trim: true
  },
  fechaActualizacion: {
    type: Date,
    default: Date.now
  }
});

// Esquema para los grupos
const grupoSchema = new mongoose.Schema({
  grupo: {
    type: String,
    required: true,
    trim: true
  },
  codigo: {
    type: String,
    required: true,
    trim: true
  },
  fechaActualizacion: {
    type: Date,
    default: Date.now
  }
});

// Crear índices para mejorar rendimiento
ordenSchema.index({ codigo: 1 });
grupoSchema.index({ grupo: 1 });
grupoSchema.index({ codigo: 1 });

// Exportar modelos
const Orden = mongoose.model('Orden', ordenSchema);
const Grupo = mongoose.model('Grupo', grupoSchema);

module.exports = { Orden, Grupo };
