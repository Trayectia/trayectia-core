// Carga las variables de entorno desde el archivo .env
require('dotenv').config();

// Dependencias principales del servidor HTTP, seguridad, archivos y correo
const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

// Garantiza disponibilidad de fetch tanto en Node nativo como vía node-fetch
const fetch = globalThis.fetch
  ? (...args) => globalThis.fetch(...args)
  : (...args) => import('node-fetch').then(({ default: f }) => f(...args));

const app = express();
const PORT = process.env.PORT || 3000;
const SERPAPI_KEY = process.env.SERPAPI_KEY;

// Avisa en consola si falta la API key de vuelos y se usará respaldo/caché
if (!SERPAPI_KEY) {
  console.warn("⚠️ AVISO: Falta SERPAPI_KEY en .env. Se usarán vuelos de respaldo si no hay clave.");
}

// Configura cabeceras de seguridad permitiendo scripts inline y recursos necesarios en la app
app.use(
  helmet({
    contentSecurityPolicy: {
      directives: {
        ...helmet.contentSecurityPolicy.getDefaultDirectives(),
        "script-src": ["'self'", "'unsafe-inline'"],
        "script-src-attr": ["'unsafe-inline'"],
        "img-src": ["'self'", "data:", "https:"],
        "connect-src": ["'self'", "https:"]
      },
    },
  })
);

// Habilita CORS, parseo de JSON y sirve los archivos estáticos desde la raíz del proyecto
app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// Limita la cantidad de peticiones a /api por IP en una ventana de 15 minutos
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 1000,
  message: { success: false, message: "Demasiadas peticiones desde esta IP, por favor intente más tarde." }
});
app.use('/api/', limiter);

// Rutas de persistencia en el directorio del proyecto
const DB_FILE = path.join(__dirname, 'trayectia_db.json');
const REPORT_FILE = path.join(__dirname, 'trayectia_reporte.xlsx');

// Lee trayectia_db.json asegurando que bookings, users y orders existan como arreglos
function loadDB() {
  if (!fs.existsSync(DB_FILE)) return { bookings: [], users: [], orders: [] };
  try {
    const parsed = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
    return {
      ...parsed,
      bookings: Array.isArray(parsed.bookings) ? parsed.bookings : [],
      users: Array.isArray(parsed.users) ? parsed.users : [],
      orders: Array.isArray(parsed.orders) ? parsed.orders : []
    };
  } catch (e) {
    return { bookings: [], users: [], orders: [] };
  }
}

// Guarda el estado actualizado de la base de datos en trayectia_db.json
function saveDB(data) {
  fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
}

// Verifica contraseñas hasheadas con formato "salt:key" usando scrypt y timingSafeEqual
function verifyPassword(password, storedHash) {
  if (!storedHash || !storedHash.includes(':')) return false;
  const [salt, key] = storedHash.split(':');
  const derived = crypto.scryptSync(password, salt, 64);
  const keyBuf = Buffer.from(key, 'hex');
  if (derived.length !== keyBuf.length) return false;
  return crypto.timingSafeEqual(derived, keyBuf);
}

// Obtiene el usuario autenticado a partir del header Authorization o req.body.authToken
function getAuthenticatedUser(req, db) {
  const rawAuth = String(req.headers?.authorization || req.headers?.Authorization || '').trim();
  let token = '';

  if (/^bearer\s+/i.test(rawAuth)) {
    token = rawAuth.replace(/^bearer\s+/i, '').trim();
  } else if (rawAuth) {
    token = rawAuth;
  }

  if (!token && req.body?.authToken) {
    token = String(req.body.authToken).trim();
  }

  if (!token || !db || !Array.isArray(db.users)) return null;

  return (
    db.users.find(
      u => Array.isArray(u.tokens) && u.tokens.some(t => String(t).trim() === token)
    ) || null
  );
}

// Calcula puntos acumulados según el monto de la compra (1 punto por cada $1.000 COP)
function calculateEarnedPoints(amount) {
  const numericAmount = Number(amount);
  if (!Number.isFinite(numericAmount) || numericAmount <= 0) return 0;
  return Math.floor(numericAmount / 1000);
}

// Parámetros financieros del programa de puntos (ajustables según tu margen real)
const POINT_REDEMPTION_VALUE_COP = 10; // Cada punto redimido equivale a $10 COP de descuento
const MIN_POINTS_TO_REDEEM = 800;      // Mínimo de puntos requeridos para activar redención ($8.000 COP mínimo)
const MAX_REDEEM_RATIO = 0.15;         // Tope máximo de redención: 15% del valor de la orden

// Calcula cuántos puntos se pueden redimir de forma segura en una orden
function calculatePointsRedemption({ requestedPoints, availablePoints, totalAmount }) {
  const grossAmount = Number(totalAmount);
  const balance = Math.max(0, Math.floor(Number(availablePoints || 0)));
  const requested = Math.max(0, Math.floor(Number(requestedPoints || 0)));

  if (
    !Number.isFinite(grossAmount) ||
    grossAmount <= 0 ||
    balance < MIN_POINTS_TO_REDEEM ||
    requested < MIN_POINTS_TO_REDEEM
  ) {
    return {
      redeemedPoints: 0,
      discountAmount: 0,
      finalAmount: Math.max(0, Number.isFinite(grossAmount) ? grossAmount : 0)
    };
  }

  const maxDiscountByOrder = Math.floor(grossAmount * MAX_REDEEM_RATIO);
  const maxPointsByOrder = Math.floor(maxDiscountByOrder / POINT_REDEMPTION_VALUE_COP);

  const redeemedPoints = Math.min(requested, balance, maxPointsByOrder);

  if (redeemedPoints < MIN_POINTS_TO_REDEEM) {
    return {
      redeemedPoints: 0,
      discountAmount: 0,
      finalAmount: Math.max(0, grossAmount)
    };
  }

  const discountAmount = redeemedPoints * POINT_REDEMPTION_VALUE_COP;
  const finalAmount = Math.max(0, grossAmount - discountAmount);

  return {
    redeemedPoints,
    discountAmount,
    finalAmount
  };
}

// Normaliza el objeto de usuario para que el frontend siempre lea points, pointsBalance y pointsHistory
function formatUserForClient(user) {
  const balance = Number(user.pointsBalance ?? user.points ?? 0);
  return {
    id: user.id,
    name: user.name || user.fullName || String(user.email || '').split('@')[0],
    email: user.email,
    points: balance,
    pointsBalance: balance,
    pointsHistory: Array.isArray(user.pointsHistory) ? user.pointsHistory : []
  };
}

// Genera un token de sesión seguro y lo asocia al usuario en trayectia_db.json
function issueAuthToken(db, user) {
  const token = crypto.randomBytes(32).toString('hex');
  user.tokens = Array.isArray(user.tokens) ? user.tokens : [];
  user.tokens.push(token);
  saveDB(db);
  return token;
}

// Configura el transporte SMTP de correo si existen las credenciales en .env
const mailTransporter = (process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS)
  ? nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: Number(process.env.SMTP_PORT || 465),
      secure: process.env.SMTP_SECURE !== 'false',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    })
  : null;

// Envía el correo de confirmación de reserva si hay transporter configurado
async function enviarCorreoReserva({ to, passengerName, pnr, amount, currency, paymentMethod }) {
  if (!to || !mailTransporter) return false;

  await mailTransporter.sendMail({
    from: process.env.MAIL_FROM || process.env.SMTP_USER,
    to,
    subject: `Confirmación de reserva Trayectia · PNR ${pnr}`,
    html: `
      <div style="font-family: Arial, sans-serif; color: #111827; max-width: 560px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px;">
        <h2 style="margin-top: 0; color: #0D9488;">¡Reserva confirmada en Trayectia!</h2>
        <p>Hola <strong>${passengerName || 'Viajero'}</strong>, hemos registrado tu solicitud de viaje.</p>
        <p><strong>Código PNR:</strong> ${pnr}</p>
        <p><strong>Monto:</strong> ${amount} ${currency || 'COP'}</p>
        <p><strong>Método de pago:</strong> ${paymentMethod || 'BINANCE'}</p>
      </div>
    `
  });

  return true;
}

// Actualiza el archivo Excel de reporte con las reservas registradas
async function updateExcelReport(dbData) {
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(REPORT_FILE)) {
    await workbook.xlsx.readFile(REPORT_FILE);
  }
  let worksheet = workbook.getWorksheet('Reservas Trayectia') || workbook.addWorksheet('Reservas Trayectia');

  worksheet.rowCount = 0;
  worksheet.columns = [
    { header: 'PNR', key: 'pnr', width: 15 },
    { header: 'Pasajero', key: 'passengerName', width: 25 },
    { header: 'Monto', key: 'amount', width: 15 },
    { header: 'Moneda', key: 'currency', width: 10 },
    { header: 'Método Pago', key: 'paymentMethod', width: 15 },
    { header: 'Puntos Ganados', key: 'pointsEarned', width: 16 },
    { header: 'Estado', key: 'status', width: 18 },
    { header: 'Fecha', key: 'createdAt', width: 20 }
  ];

  dbData.bookings.forEach(b => worksheet.addRow(b));
  await workbook.xlsx.writeFile(REPORT_FILE);
}

// Catálogo base de aeropuertos para sugerencias de origen y destino
const AIRPORTS_CATALOG = [
  { iata: 'BOG', city: 'Bogotá', name: 'El Dorado International Airport', country: 'Colombia' },
  { iata: 'CTG', city: 'Cartagena', name: 'Rafael Núñez International Airport', country: 'Colombia' },
  { iata: 'MDE', city: 'Medellín', name: 'José María Córdova International Airport', country: 'Colombia' },
  { iata: 'CLO', city: 'Cali', name: 'Alfonso Bonilla Aragón International Airport', country: 'Colombia' },
  { iata: 'BAQ', city: 'Barranquilla', name: 'Ernesto Cortissoz International Airport', country: 'Colombia' },
  { iata: 'SMR', city: 'Santa Marta', name: 'Simón Bolívar International Airport', country: 'Colombia' },
  { iata: 'ADZ', city: 'San Andrés', name: 'Gustavo Rojas Pinilla International Airport', country: 'Colombia' },
  { iata: 'PEI', city: 'Pereira', name: 'Matecaña International Airport', country: 'Colombia' },
  { iata: 'BGA', city: 'Bucaramanga', name: 'Palonegro International Airport', country: 'Colombia' },
  { iata: 'CUC', city: 'Cúcuta', name: 'Camilo Daza International Airport', country: 'Colombia' },
  { iata: 'GRU', city: 'São Paulo', name: 'Guarulhos International Airport', country: 'Brasil' },
  { iata: 'CGH', city: 'São Paulo', name: 'Congonhas Airport', country: 'Brasil' },
  { iata: 'GIG', city: 'Río de Janeiro', name: 'Galeão International Airport', country: 'Brasil' },
  { iata: 'SDU', city: 'Río de Janeiro', name: 'Santos Dumont Airport', country: 'Brasil' },
  { iata: 'MIA', city: 'Miami', name: 'Miami International Airport', country: 'Estados Unidos' },
  { iata: 'FLL', city: 'Fort Lauderdale', name: 'Fort Lauderdale-Hollywood International Airport', country: 'Estados Unidos' },
  { iata: 'JFK', city: 'Nueva York', name: 'John F. Kennedy International Airport', country: 'Estados Unidos' },
  { iata: 'MAD', city: 'Madrid', name: 'Adolfo Suárez Madrid-Barajas Airport', country: 'España' },
  { iata: 'LIS', city: 'Lisboa', name: 'Humberto Delgado Airport', country: 'Portugal' },
  { iata: 'SCL', city: 'Santiago de Chile', name: 'Arturo Merino Benítez International Airport', country: 'Chile' },
  { iata: 'EZE', city: 'Buenos Aires', name: 'Ministro Pistarini International Airport', country: 'Argentina' },
  { iata: 'LIM', city: 'Lima', name: 'Jorge Chávez International Airport', country: 'Perú' },
  { iata: 'PTY', city: 'Ciudad de Panamá', name: 'Tocumen International Airport', country: 'Panamá' },
  { iata: 'LHR', city: 'Londres', name: 'Heathrow Airport', country: 'Reino Unido' }
];

// Normaliza texto en minúsculas sin tildes para búsquedas y filtros
function normalizeSearchText(text = '') {
  return text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

// Endpoint de autocompletado/sugerencias de aeropuertos y ciudades
app.get('/api/places/suggest', (req, res) => {
  const q = normalizeSearchText(req.query.q || '');
  if (!q) {
    return res.json({ success: true, places: AIRPORTS_CATALOG.slice(0, 8) });
  }

  const matches = AIRPORTS_CATALOG.filter((item) => {
    const haystack = normalizeSearchText(`${item.iata} ${item.city} ${item.name} ${item.country}`);
    return haystack.includes(q);
  }).slice(0, 8);

  res.json({ success: true, places: matches });
});

// Configuración de caché en disco para búsquedas de vuelos (TTL de 2 horas)
const FLIGHT_CACHE_FILE = path.join(__dirname, 'flights_cache.json');
const CACHE_TTL_MS = 2 * 60 * 60 * 1000;

function loadFlightCache() {
  if (!fs.existsSync(FLIGHT_CACHE_FILE)) return {};
  try {
    const raw = JSON.parse(fs.readFileSync(FLIGHT_CACHE_FILE, 'utf8'));
    return raw && typeof raw === 'object' ? raw : {};
  } catch (_) {
    return {};
  }
}

function saveFlightCache(cacheObj) {
  try {
    fs.writeFileSync(FLIGHT_CACHE_FILE, JSON.stringify(cacheObj, null, 2));
  } catch (err) {
    console.error('Error guardando caché de vuelos:', err.message);
  }
}

// Interpreta la cantidad de pasajeros por tipo desde query params o JSON
function parsePassengersBreakdown(passengersParam, query = {}) {
  let adults = Number(query.adults ?? 0);
  let children = Number(query.children ?? 0);
  let infantsInSeat = Number(query.infants_in_seat ?? query.infantsInSeat ?? 0);
  let infantsOnLap = Number(query.infants_on_lap ?? query.infantsOnLap ?? 0);

  if (passengersParam) {
    try {
      const parsed = typeof passengersParam === 'string' ? JSON.parse(passengersParam) : passengersParam;
      if (Array.isArray(parsed) && parsed.length > 0) {
        adults = parsed.filter(p => p.type === 'adult').length;
        children = parsed.filter(p => p.type === 'child').length;
        infantsInSeat = parsed.filter(p => p.type === 'infant_in_seat').length;
        infantsOnLap = parsed.filter(p => p.type === 'infant_on_lap' || p.type === 'infant').length;
      } else if (parsed && typeof parsed === 'object') {
        adults = Number(parsed.adults ?? adults);
        children = Number(parsed.children ?? children);
        infantsInSeat = Number(parsed.infants_in_seat ?? parsed.infantsInSeat ?? infantsInSeat);
        infantsOnLap = Number(parsed.infants_on_lap ?? parsed.infantsOnLap ?? infantsOnLap);
      }
    } catch (_) {}
  }

  return {
    adults: Math.max(1, Number.isFinite(adults) && adults > 0 ? Math.floor(adults) : 1),
    children: Math.max(0, Number.isFinite(children) ? Math.floor(children) : 0),
    infantsInSeat: Math.max(0, Number.isFinite(infantsInSeat) ? Math.floor(infantsInSeat) : 0),
    infantsOnLap: Math.max(0, Number.isFinite(infantsOnLap) ? Math.floor(infantsOnLap) : 0)
  };
}

// Normaliza el tipo de viaje ("oneway" o "roundtrip")
function normalizeTripType(rawType, returnDate) {
  const val = String(rawType || '').toLowerCase().trim();
  if (val === 'oneway' || val === 'one_way' || val === 'ow' || val === '2') return 'oneway';
  if (val === 'roundtrip' || val === 'round_trip' || val === 'rt' || val === '1') return 'roundtrip';
  return returnDate ? 'roundtrip' : 'oneway';
}

// Normaliza el código de clase de cabina (1: Economy, 2: Premium Economy, 3: Business, 4: First)
function normalizeTravelClass(rawClass) {
  const val = String(rawClass || '1').toLowerCase().trim();
  const map = {
    '1': '1', 'economy': '1', 'economica': '1', 'económica': '1',
    '2': '2', 'premium_economy': '2', 'premium': '2',
    '3': '3', 'business': '3', 'ejecutiva': '3',
    '4': '4', 'first': '4', 'primera': '4'
  };
  return map[val] || '1';
}

// Convierte el código de cabina a etiqueta legible
function getCabinLabel(travelClassCode) {
  const labels = {
    '1': 'Economy',
    '2': 'Premium Economy',
    '3': 'Business',
    '4': 'First'
  };
  return labels[String(travelClassCode || '1')] || 'Economy';
}

// Normaliza fechas/horas de Google Flights evitando fechas duplicadas
function normalizeFlightTimestamp(rawTime, fallbackDate = '') {
  const cleaned = String(rawTime || '')
    .trim()
    .replace(/^(\d{4}-\d{2}-\d{2}T)\1/, '$1');

  if (!cleaned) {
    return fallbackDate ? `${fallbackDate}T08:00:00` : '';
  }

  if (/^\d{4}-\d{2}-\d{2}[T\s]\d{2}:\d{2}/.test(cleaned)) {
    const normalized = cleaned.replace(' ', 'T');
    return normalized.length === 16 ? `${normalized}:00` : normalized;
  }

  if (/^\d{1,2}:\d{2}$/.test(cleaned)) {
    const [hh, mm] = cleaned.split(':');
    const timeStr = `${hh.padStart(2, '0')}:${mm}:00`;
    return fallbackDate ? `${fallbackDate}T${timeStr}` : timeStr;
  }

  return cleaned;
}

// Consulta vuelos en SerpAPI (Google Flights) en pesos colombianos (COP)
async function requestGoogleFlights({
  origin,
  destination,
  departure_date,
  return_date,
  tripType = 'oneway',
  pax = { adults: 1, children: 0, infantsInSeat: 0, infantsOnLap: 0 },
  travelClass = '1',
  stops = '',
  maxPrice = ''
}) {
  const isRoundTrip = tripType === 'roundtrip' && Boolean(return_date);
  const params = new URLSearchParams({
    engine: 'google_flights',
    departure_id: origin,
    arrival_id: destination,
    outbound_date: departure_date,
    currency: 'COP',
    hl: 'es',
    adults: String(pax.adults),
    children: String(pax.children),
    infants_in_seat: String(pax.infantsInSeat),
    infants_on_lap: String(pax.infantsOnLap),
    travel_class: String(travelClass || '1'),
    type: isRoundTrip ? '1' : '2',
    api_key: SERPAPI_KEY
  });

  if (isRoundTrip && return_date) params.set('return_date', return_date);
  if (stops !== '' && stops !== undefined) params.set('stops', String(stops));
  if (maxPrice && Number(maxPrice) > 0) params.set('max_price', String(Math.floor(Number(maxPrice))));

  const response = await fetch(`https://serpapi.com/search.json?${params.toString()}`);
  return response.json();
}

// Transforma la respuesta de Google Flights a un formato uniforme para el frontend
function formatGoogleFlightsList(data = {}, departureDate = '', cabinLabel = 'Economy') {
  const rawList = [
    ...(Array.isArray(data.best_flights) ? data.best_flights : []),
    ...(Array.isArray(data.other_flights) ? data.other_flights : [])
  ];

  return rawList
    .filter(item => Array.isArray(item.flights) && item.flights.length > 0 && item.price)
    .map((item, idx) => {
      const firstLeg = item.flights[0];
      const lastLeg = item.flights[item.flights.length - 1];
      const departureTime = normalizeFlightTimestamp(firstLeg.departure_airport?.time, departureDate);
      const arrivalTime = normalizeFlightTimestamp(lastLeg.arrival_airport?.time, departureDate);

      return {
        id: `gf_${Date.now()}_${idx}`,
        airline: firstLeg.airline || 'Aerolínea',
        flightNumber: firstLeg.flight_number || 'GF',
        origin: firstLeg.departure_airport?.id || '',
        destination: lastLeg.arrival_airport?.id || '',
        departureTime,
        arrivalTime,
        durationMinutes: Number(item.total_duration || 120),
        stops: Math.max(0, item.flights.length - 1),
        cabin: cabinLabel,
        price: Number(item.price),
        currency: 'COP'
      };
    });
}

// Aplica filtros secundarios de escalas, aerolínea, precio máximo, franja horaria y duración
function applySecondaryFlightFilters(flights = [], filters = {}) {
  const {
    maxStops = '',
    airline = '',
    maxPrice = '',
    timeWindow = '',
    maxDurationMinutes = ''
  } = filters;

  return flights.filter((f) => {
    if (maxStops !== '' && maxStops !== undefined && maxStops !== null) {
      const stopLimit = Number(maxStops);
      if (Number.isFinite(stopLimit) && Number(f.stops || 0) > stopLimit) return false;
    }
    if (airline && String(airline).toLowerCase() !== 'all') {
      if (!normalizeSearchText(String(f.airline || '')).includes(normalizeSearchText(String(airline)))) {
        return false;
      }
    }
    if (maxPrice && Number(maxPrice) > 0 && Number(f.price || 0) > Number(maxPrice)) return false;
    if (maxDurationMinutes && Number(maxDurationMinutes) > 0 && Number(f.durationMinutes || 0) > Number(maxDurationMinutes)) {
      return false;
    }
    if (timeWindow && String(timeWindow).toLowerCase() !== 'all') {
      const depStr = String(f.departureTime || '');
      const timePart = depStr.includes('T') ? depStr.split('T')[1] : (depStr.split(' ')[1] || '');
      const hour = parseInt(timePart.slice(0, 2), 10);
      if (Number.isFinite(hour)) {
        if (timeWindow === 'early_morning' && !(hour >= 0 && hour < 6)) return false;
        if (timeWindow === 'morning' && !(hour >= 6 && hour < 12)) return false;
        if (timeWindow === 'afternoon' && !(hour >= 12 && hour < 18)) return false;
        if (timeWindow === 'night' && !(hour >= 18 && hour <= 23)) return false;
      }
    }
    return true;
  });
}

const MOCK_FLIGHTS = String(process.env.MOCK_FLIGHTS || '').toLowerCase() === 'true';

// Endpoint principal de búsqueda de vuelos con caché (sin vuelos falsos de respaldo)
app.get('/api/flights/search', async (req, res) => {
  try {
    const origin = (req.query.origin || '').toUpperCase().trim();
    const destination = (req.query.destination || '').toUpperCase().trim();
    const departure_date = (req.query.departure_date || req.query.date || '').trim();
    const rawReturnDate = (req.query.returnDate || req.query.return_date || '').trim();
    const tripType = normalizeTripType(req.query.tripType || req.query.type, rawReturnDate);
    const return_date = tripType === 'roundtrip' ? rawReturnDate : '';
    const pax = parsePassengersBreakdown(req.query.passengers, req.query);
    const travelClass = normalizeTravelClass(req.query.travelClass || req.query.travel_class || req.query.cabin);
    const cabinLabel = getCabinLabel(travelClass);
    const stops = (req.query.stops ?? '').toString().trim();
    const maxPrice = (req.query.maxPrice || req.query.max_price || '').toString().trim();
    const airlineFilter = (req.query.airline || '').toString().trim();
    const timeWindow = (req.query.timeWindow || req.query.departureWindow || '').toString().trim();
    const maxDurationMinutes = (req.query.maxDurationMinutes || req.query.max_duration || '').toString().trim();

    if (!origin || !destination || !departure_date) {
      return res.status(400).json({
        success: false,
        message: 'Faltan parámetros de búsqueda (origin, destination, date o departure_date).'
      });
    }

    const paxKey = `${pax.adults}A-${pax.children}C-${pax.infantsInSeat}IS-${pax.infantsOnLap}IL`;
    const cacheKey = `${origin}-${destination}-${departure_date}-${tripType === 'roundtrip' ? (return_date || 'RT') : 'OW'}-${paxKey}-C${travelClass}-S${stops || 'ALL'}`;
    const diskCache = loadFlightCache();
    const cached = diskCache[cacheKey];

    let formattedFlights = [];
    let offerRequestId = "gf-" + Date.now();
    let fromCache = false;

    if (cached && (Date.now() - Number(cached.timestamp || 0) < CACHE_TTL_MS)) {
      formattedFlights = Array.isArray(cached.flights) ? cached.flights : [];
      offerRequestId = cached.offerRequestId || ('cache-' + Date.now());
      fromCache = true;
    } else {
      if (!MOCK_FLIGHTS && SERPAPI_KEY) {
        const gfData = await requestGoogleFlights({
          origin, destination, departure_date, return_date, tripType, pax, travelClass, stops, maxPrice
        });
        if (gfData && gfData.error) {
          return res.status(400).json({ success: false, message: gfData.error || "Error consultando Google Flights" });
        }
        formattedFlights = formatGoogleFlightsList(gfData, departure_date, cabinLabel);
      }

      if (formattedFlights.length > 0) {
        diskCache[cacheKey] = {
          timestamp: Date.now(),
          flights: formattedFlights,
          offerRequestId,
          usedOneWayFallback: false
        };
        saveFlightCache(diskCache);
      }
    }

    const filteredFlights = applySecondaryFlightFilters(formattedFlights, {
      maxStops: stops,
      airline: airlineFilter,
      maxPrice,
      timeWindow,
      maxDurationMinutes
    });

    return res.json({
      success: true,
      flights: filteredFlights,
      totalBeforeFilters: formattedFlights.length,
      tripType,
      passengersBreakdown: pax,
      travelClass,
      cabin: cabinLabel,
      offerRequestId,
      usedOneWayFallback: false,
      fromCache,
      mockMode: MOCK_FLIGHTS
    });
  } catch (error) {
    console.error("Error en GET /api/flights/search:", error);
    return res.status(500).json({ success: false, message: "Error consultando vuelos en el servidor." });
  }
});

// Registra un nuevo usuario con contraseña hasheada (formato salt:hex)
app.post('/api/auth/register', (req, res) => {
  try {
    const name = String(req.body?.name || req.body?.fullName || '').trim();
    const email = String(req.body?.email || '').trim().toLowerCase();
    const password = String(req.body?.password || '');

    if (!email || !password) {
      return res.status(400).json({
        ok: false,
        success: false,
        message: 'Ingresa tu correo electrónico y contraseña.'
      });
    }

    const db = loadDB();
    db.users = Array.isArray(db.users) ? db.users : [];

    const existing = db.users.find(
      (u) => String(u.email || '').trim().toLowerCase() === email
    );
    if (existing) {
      return res.status(409).json({
        ok: false,
        success: false,
        message: 'Ya existe una cuenta registrada con ese correo.'
      });
    }

    const salt = crypto.randomBytes(16).toString('hex');
    const derivedKey = crypto.scryptSync(password, salt, 64).toString('hex');
    const passwordHash = `${salt}:${derivedKey}`;

    const newUser = {
      id: `usr_${Date.now()}`,
      name: name || email.split('@')[0],
      email,
      passwordHash,
      points: 0,
      pointsBalance: 0,
      pointsHistory: [],
      tokens: [],
      createdAt: new Date().toISOString()
    };

    db.users.push(newUser);
    const token = issueAuthToken(db, newUser);

    return res.status(201).json({
      ok: true,
      success: true,
      token,
      user: formatUserForClient(newUser)
    });
  } catch (err) {
    console.error('Error en POST /api/auth/register:', err);
    return res.status(500).json({
      ok: false,
      success: false,
      message: 'Error interno al registrar el usuario.'
    });
  }
});

// Inicia sesión validando correo y contraseña y devuelve token + saldo de puntos
app.post('/api/auth/login', (req, res) => {
  try {
    const email = String(req.body?.email || '').trim().toLowerCase();
    const password = String(req.body?.password || '');

    if (!email || !password) {
      return res.status(400).json({
        ok: false,
        success: false,
        message: 'Ingresa correo y contraseña para iniciar sesión.'
      });
    }

    const db = loadDB();
    db.users = Array.isArray(db.users) ? db.users : [];

    const user = db.users.find(
      (u) => String(u.email || '').trim().toLowerCase() === email
    );

    const storedHash = user?.passwordHash || user?.password || '';
    const validPass =
      user &&
      ((storedHash.includes(':') && verifyPassword(password, storedHash)) ||
        (!storedHash.includes(':') && storedHash && storedHash === password));

    if (!user || !validPass) {
      return res.status(401).json({
        ok: false,
        success: false,
        message: 'Correo o contraseña incorrectos.'
      });
    }

    const token = issueAuthToken(db, user);

    return res.json({
      ok: true,
      success: true,
      token,
      user: formatUserForClient(user)
    });
  } catch (err) {
    console.error('Error en POST /api/auth/login:', err);
    return res.status(500).json({
      ok: false,
      success: false,
      message: 'Error interno al iniciar sesión.'
    });
  }
});

// Devuelve el perfil y saldo actualizado del usuario autenticado por token
app.get('/api/auth/me', (req, res) => {
  const db = loadDB();
  const user = getAuthenticatedUser(req, db);
  if (!user) {
    return res.status(401).json({
      ok: false,
      success: false,
      message: 'Sesión no válida o expirada.'
    });
  }
  return res.json({
    ok: true,
    success: true,
    user: formatUserForClient(user)
  });
});

// Evalúa si la hora actual en Colombia está dentro del horario operativo (Lun-Sáb 08:00 - 20:00 COT)
function getColombiaOperatingStatus(now = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Bogota',
    weekday: 'short',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  }).formatToParts(now);

  const weekday = parts.find((p) => p.type === 'weekday')?.value || 'Mon';
  const hour = parseInt(parts.find((p) => p.type === 'hour')?.value || '0', 10);
  const minute = parseInt(parts.find((p) => p.type === 'minute')?.value || '0', 10);
  const totalMinutes = hour * 60 + minute;

  const isWorkDay = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].includes(weekday);
  const isOpen = isWorkDay && totalMinutes >= 8 * 60 && totalMinutes <= 20 * 60;

  return {
    isOpen,
    scheduleWindow: 'Lun-Sáb 08:00 - 20:00 (hora Colombia)',
    statusWhenPaid: isOpen ? 'paid_pending_issue' : 'pending_schedule',
    operatingMessage: isOpen
      ? 'Dentro del horario operativo de Trayectia: una vez confirmado el pago, la orden pasa a paid_pending_issue para emisión.'
      : 'Fuera del horario operativo de Trayectia (Lun-Sáb 08:00 - 20:00 COT): quedará en pending_schedule para emisión en la siguiente ventana operativa.'
  };
}

// Genera un código único de orden con prefijo TRX- (ej. TRX-BR493W)
function generateOrderCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let suffix = '';
  for (let i = 0; i < 6; i++) {
    suffix += chars[crypto.randomInt(0, chars.length)];
  }
  return `TRX-${suffix}`;
}

// Crea y persiste una orden aplicando redención segura (mínimo 800 pts) y acumulación (1 pt/$1.000)
app.post('/api/orders', (req, res) => {
  try {
    const {
      paymentMethod = 'ADDI',
      flight = {},
      contact = {},
      passengers = [],
      pricing = {},
      paymentApproved = false
    } = req.body || {};

    const normalizedMethod = String(paymentMethod || 'ADDI').toUpperCase();
    const email = String(contact.email || '').trim().toLowerCase();
    const phone = String(contact.phone || '').trim();

    if (!email || !Array.isArray(passengers) || passengers.length === 0) {
      return res.status(400).json({
        ok: false,
        success: false,
        message: 'Faltan datos de contacto (correo) o pasajeros.'
      });
    }

    const db = loadDB();
    const authUser = getAuthenticatedUser(req, db);
    const matchedUser =
      authUser ||
      db.users.find((u) => String(u.email || '').trim().toLowerCase() === email) ||
      null;

    const orderCode = generateOrderCode();
    const schedule = getColombiaOperatingStatus(new Date());
    const status = paymentApproved ? schedule.statusWhenPaid : 'pending_payment';

    const grossAmount = Number(pricing?.grossAmount ?? pricing?.totalAmount ?? pricing?.amount ?? 0);
    const requestedPoints = Number(pricing?.pointsToRedeem ?? req.body?.pointsToRedeem ?? 0);
    const currentUserPoints = matchedUser
      ? Number(matchedUser.pointsBalance ?? matchedUser.points ?? 0)
      : 0;

    const redemption = matchedUser
      ? calculatePointsRedemption({
          requestedPoints,
          availablePoints: currentUserPoints,
          totalAmount: grossAmount
        })
      : {
          redeemedPoints: 0,
          discountAmount: 0,
          finalAmount: Math.max(0, grossAmount)
        };

    const earnedPoints = calculateEarnedPoints(redemption.finalAmount);

    const paymentReference =
      normalizedMethod === 'CRYPTO' || normalizedMethod === 'BINANCE'
        ? `BIN-${orderCode}`
        : `PAY-${orderCode}`;

    const newOrder = {
      orderCode,
      paymentReference,
      paymentMethod: normalizedMethod,
      status,
      userId: matchedUser ? matchedUser.id : null,
      pointsRedeemed: redemption.redeemedPoints,
      discountAmount: redemption.discountAmount,
      pointsEarned: earnedPoints,
      operatingSchedule: {
        isOpen: schedule.isOpen,
        window: schedule.scheduleWindow
      },
      operatingMessage: schedule.operatingMessage,
      flight: {
        origin: String(flight.origin || '').trim().toUpperCase(),
        destination: String(flight.destination || '').trim().toUpperCase(),
        departureDate: flight.departureDate || '',
        returnDate: flight.returnDate || '',
        details: flight.details || null
      },
      contact: {
        email,
        phone
      },
      passengers,
      pricing: {
        ...pricing,
        grossAmount,
        discountAmount: redemption.discountAmount,
        pointsRedeemed: redemption.redeemedPoints,
        finalAmount: redemption.finalAmount,
        totalAmount: redemption.finalAmount
      },
      createdAt: new Date().toISOString()
    };

    db.orders = Array.isArray(db.orders) ? db.orders : [];
    db.orders.push(newOrder);

    if (matchedUser && (redemption.redeemedPoints > 0 || earnedPoints > 0)) {
      const nextBal = currentUserPoints - redemption.redeemedPoints + earnedPoints;
      matchedUser.points = nextBal;
      matchedUser.pointsBalance = nextBal;
      matchedUser.pointsHistory = Array.isArray(matchedUser.pointsHistory)
        ? matchedUser.pointsHistory
        : [];

      const routeLabel = `${newOrder.flight.origin || '---'} → ${newOrder.flight.destination || '---'}`;

      if (redemption.redeemedPoints > 0) {
        matchedUser.pointsHistory.push({
          pnr: orderCode,
          points: -redemption.redeemedPoints,
          description: `Redención en orden ${orderCode} (-$${redemption.discountAmount.toLocaleString()} COP · ${routeLabel})`,
          createdAt: new Date().toISOString()
        });
      }

      if (earnedPoints > 0) {
        matchedUser.pointsHistory.push({
          pnr: orderCode,
          points: earnedPoints,
          description: `Acumulación por orden ${orderCode} (${routeLabel})`,
          createdAt: new Date().toISOString()
        });
      }
    }

    saveDB(db);

    return res.status(201).json({
      ok: true,
      success: true,
      data: newOrder,
      user: matchedUser ? formatUserForClient(matchedUser) : null
    });
  } catch (error) {
    console.error('Error en POST /api/orders:', error);
    return res.status(500).json({
      ok: false,
      success: false,
      message: 'Error interno al crear la orden.'
    });
  }
});

// Lista las reservas y órdenes asociadas al usuario autenticado
app.get('/api/my-bookings', (req, res) => {
  const db = loadDB();
  const user = getAuthenticatedUser(req, db);
  if (!user) {
    return res.status(401).json({
      ok: false,
      success: false,
      message: 'Debes iniciar sesión para ver tus reservas.'
    });
  }

  const userEmail = String(user.email || '').trim().toLowerCase();
  const myBookings = (db.bookings || []).filter(
    (b) => b.userId === user.id || String(b.contact?.email || b.email || '').trim().toLowerCase() === userEmail
  );
  const myOrders = (db.orders || []).filter(
    (o) => o.userId === user.id || String(o.contact?.email || o.email || '').trim().toLowerCase() === userEmail
  );

  return res.json({
    ok: true,
    success: true,
    bookings: myBookings,
    orders: myOrders
  });
});

// Consulta una reserva o una orden en trayectia_db.json por PNR, orderCode, code o id
app.get('/api/bookings/lookup', (req, res) => {
  try {
    const rawCode = String(
      req.query.pnr || req.query.code || req.query.orderCode || ''
    )
      .trim()
      .toUpperCase();

    if (!rawCode) {
      return res.status(400).json({
        success: false,
        ok: false,
        message: 'Ingresa un código PNR o de orden válido.'
      });
    }

    const db = loadDB();
    const bookingsList = Array.isArray(db.bookings) ? db.bookings : [];
    const ordersList = Array.isArray(db.orders) ? db.orders : [];

    const matchesCode = (item) => {
      if (!item || typeof item !== 'object') return false;
      const candidates = [
        item.pnr,
        item.orderCode,
        item.code,
        item.id,
        item.bookingCode,
        item.paymentReference
      ]
        .filter(Boolean)
        .map((val) => String(val).trim().toUpperCase());

      return candidates.includes(rawCode);
    };

    const found = bookingsList.find(matchesCode) || ordersList.find(matchesCode);

    if (!found) {
      return res.status(404).json({
        success: false,
        ok: false,
        message: `No se encontró ninguna reserva u orden con el código ${rawCode}.`
      });
    }

    const firstPax =
      Array.isArray(found.passengers) && found.passengers.length > 0
        ? found.passengers[0]
        : null;

    const passengerName =
      found.passengerName ||
      found.passenger ||
      (firstPax
        ? `${firstPax.givenName || ''} ${firstPax.familyName || ''}`.trim()
        : 'Viajero');

    const details = found.flight?.details || found.flight || {};
    const grossAmount = Number(
      found.pricing?.grossAmount ??
      found.grossAmount ??
      found.pricing?.totalAmount ??
      found.amount ??
      0
    );
    const discountAmount = Number(
      found.discountAmount ??
      found.pricing?.discountAmount ??
      0
    );
    const finalAmount = Number(
      found.pricing?.finalAmount ??
      found.finalAmount ??
      found.pricing?.totalAmount ??
      found.amount ??
      0
    );
    const pointsRedeemed = Number(
      found.pointsRedeemed ??
      found.pricing?.pointsRedeemed ??
      0
    );
    const pointsEarned = Number(
      found.pointsEarned ?? calculateEarnedPoints(finalAmount)
    );

    const resolvedCode = String(
      found.orderCode || found.pnr || found.id || rawCode
    )
      .trim()
      .toUpperCase();

    const bookingPayload = {
      pnr: resolvedCode,
      orderCode: resolvedCode,
      paymentReference: found.paymentReference || '',
      passengerName,
      contactEmail: found.contact?.email || found.email || '',
      contactPhone: found.contact?.phone || found.phone || '',
      status: found.status || 'pending_payment',
      origin: String(found.origin || details.origin || found.flight?.origin || '').trim().toUpperCase(),
      destination: String(found.destination || details.destination || found.flight?.destination || '').trim().toUpperCase(),
      airline: found.airline || details.airline || '',
      flightNumber: found.flightNumber || details.flightNumber || '',
      departureTime: found.departureTime || details.departureTime || found.flight?.departureDate || '',
      arrivalTime: found.arrivalTime || details.arrivalTime || found.flight?.returnDate || '',
      grossAmount: Number.isFinite(grossAmount) ? grossAmount : 0,
      discountAmount: Number.isFinite(discountAmount) ? discountAmount : 0,
      finalAmount: Number.isFinite(finalAmount) ? finalAmount : 0,
      amount: Number.isFinite(finalAmount) ? finalAmount : 0,
      currency: found.currency || found.pricing?.currency || 'COP',
      paymentMethod: found.paymentMethod || 'ADDI',
      pointsRedeemed: Number.isFinite(pointsRedeemed) ? pointsRedeemed : 0,
      pointsEarned: Number.isFinite(pointsEarned) ? pointsEarned : 0,
      operatingMessage: found.operatingMessage || '',
      createdAt: found.createdAt || '',
      emailSent: Boolean(found.emailSent)
    };

    return res.json({
      success: true,
      ok: true,
      booking: bookingPayload,
      data: bookingPayload
    });
  } catch (err) {
    console.error('Error en GET /api/bookings/lookup:', err);
    return res.status(500).json({
      success: false,
      ok: false,
      message: 'Error interno al consultar la reserva.'
    });
  }
});

// Devuelve destinos destacados para las tarjetas de ofertas y evita el 404 en frontend
app.get('/api/destinos-ofertas', (req, res) => {
  const roundTrip = String(req.query.roundTrip || 'true').toLowerCase() === 'true';
  const diskCache = loadFlightCache();
  const cachedValues = Object.values(diskCache || {});

  const offers = AIRPORTS_CATALOG
    .filter((item) => item.iata !== 'BOG')
    .slice(0, 8)
    .map((dest) => {
      const match = cachedValues.find(
        (entry) => Array.isArray(entry?.flights) && entry.flights.some((f) => f.destination === dest.iata)
      );
      const minPrice = match
        ? Math.min(...match.flights.map((f) => Number(f.price || 0)).filter((p) => p > 0))
        : null;

      return {
        origin: 'BOG',
        destination: dest.iata,
        city: dest.city,
        country: dest.country,
        roundTrip,
        minPrice: Number.isFinite(minPrice) ? minPrice : null,
        currency: 'COP'
      };
    });

  return res.json({
    ok: true,
    success: true,
    roundTrip,
    destinations: offers
  });
});

// Sirve la página principal index.html en la ruta raíz
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Arranca el servidor Express en el puerto configurado
app.listen(PORT, () => {
  console.log(`Servidor Trayectia corriendo en http://localhost:${PORT}`);
});
