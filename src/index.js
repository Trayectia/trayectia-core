const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

const DB_PATH = path.join(__dirname, '..', 'trayectia_db.json');
const PUBLIC_DIR = path.join(__dirname, '..', 'public');

const TASA_ADDI = 0.07735;
const COSTO_OPERATIVO = 8000;
const APORTE_FUNDACION_PCT = 0.15;
const TOLERANCIA_PCT = 0.20;
const TOLERANCIA_TOPE = 5000;

app.use(cors());
app.use(express.json());
app.use(express.static(PUBLIC_DIR));

function leerDB() {
  if (!fs.existsSync(DB_PATH)) {
    const inicial = { reservas: [] };
    fs.writeFileSync(DB_PATH, JSON.stringify(inicial, null, 2));
    return inicial;
  }
  return JSON.parse(fs.readFileSync(DB_PATH, 'utf8'));
}

function guardarDB(data) {
  fs.writeFileSync(DB_PATH, JSON.stringify(data, null, 2));
}

function calcularFinanzas(tarifaBase, usarAddi = true) {
  const base = Number(tarifaBase) || 0;
  const fee = Math.min(99000, Math.round(25000 + base * 0.015));
  const subtotalAgencia = base + fee;
  const recargoAddi = usarAddi ? Math.round(subtotalAgencia * TASA_ADDI) : 0;
  const totalCliente = subtotalAgencia + recargoAddi;

  const utilidadPreviaBruta = fee;
  const aporteFundacion = Math.round(Math.max(0, utilidadPreviaBruta) * APORTE_FUNDACION_PCT);
  const utilidadNetaEstimada = utilidadPreviaBruta - COSTO_OPERATIVO - aporteFundacion;

  const toleranciaPermitida = Math.min(
    TOLERANCIA_TOPE,
    Math.round(Math.max(0, utilidadPreviaBruta) * TOLERANCIA_PCT)
  );

  return {
    tarifaBase: base,
    fee,
    usarAddi,
    recargoAddi,
    totalCliente,
    utilidadPreviaBruta,
    costoOperativo: COSTO_OPERATIVO,
    aporteFundacion,
    utilidadNetaEstimada,
    toleranciaPermitida
  };
}

function generarCodigoTRX() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = 'TRX-';
  for (let i = 0; i < 6; i++) {
    code += chars[Math.floor(Math.random() * chars.length)];
  }
  return code;
}

app.get('/api/reservas', (req, res) => {
  const db = leerDB();
  res.json(db.reservas);
});

app.post('/api/reservas', (req, res) => {
  const { pasajero, ruta, tarifaBase, usarAddi = true } = req.body;
  if (!pasajero || !ruta || !Number(tarifaBase)) {
    return res.status(400).json({ error: 'Faltan datos: pasajero, ruta o tarifaBase válida.' });
  }

  const db = leerDB();
  const finanzas = calcularFinanzas(tarifaBase, Boolean(usarAddi));
  const reserva = {
    pnr: generarCodigoTRX(),
    pasajero,
    ruta,
    estado: 'PAGADO_POR_EMITIR',
    creadoEn: new Date().toISOString(),
    snapshot: finanzas,
    tarifaActualValidada: finanzas.tarifaBase,
    codigoAerolinea: null
  };

  db.reservas.unshift(reserva);
  guardarDB(db);
  res.status(201).json(reserva);
});

app.post('/api/reservas/:pnr/validar-emision', (req, res) => {
  const { tarifaActual } = req.body;
  const db = leerDB();
  const reserva = db.reservas.find((r) => r.pnr === req.params.pnr);
  if (!reserva) return res.status(404).json({ error: 'Reserva no encontrada.' });

  const actual = Number(tarifaActual);
  const baseInicial = reserva.snapshot.tarifaBase;
  const variacionSubida = Math.max(0, actual - baseInicial);
  reserva.tarifaActualValidada = actual;

  if (variacionSubida > reserva.snapshot.toleranciaPermitida) {
    reserva.estado = 'ALERTA_VARIACION_TARIFA';
  } else {
    reserva.estado = 'LISTA_PARA_EMITIR';
  }

  guardarDB(db);
  res.json({
    reserva,
    variacionSubida,
    toleranciaPermitida: reserva.snapshot.toleranciaPermitida
  });
});

app.post('/api/reservas/:pnr/emitir', (req, res) => {
  const { codigoAerolinea } = req.body;
  const db = leerDB();
  const reserva = db.reservas.find((r) => r.pnr === req.params.pnr);
  if (!reserva) return res.status(404).json({ error: 'Reserva no encontrada.' });

  if (reserva.estado === 'ALERTA_VARIACION_TARIFA') {
    return res.status(409).json({
      error: 'Reserva en ALERTA_VARIACION_TARIFA. Resuelve la variación antes de emitir.'
    });
  }

  reserva.estado = 'EMITIDO';
  reserva.codigoAerolinea = codigoAerolinea || `PNR-${Math.floor(100000 + Math.random() * 900000)}`;
  reserva.emitidoEn = new Date().toISOString();

  guardarDB(db);
  res.json(reserva);
});

app.get('*', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Servidor Trayectia activo en puerto ${PORT}`);
});
