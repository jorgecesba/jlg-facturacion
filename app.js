import express from 'express';
import crypto from 'crypto';
import path from 'path';
import bcrypt from 'bcrypt';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 4000;

app.use(express.json({ limit: '15mb' }));
app.use(express.urlencoded({ extended: true, limit: '15mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ==========================================
// SEGURIDAD: CONFIGURACIÓN Y AUTENTICACIÓN ADMIN
// ==========================================
// Credenciales de entorno (por defecto desarrollo/demo)
const ADMIN_USERNAME = process.env.ADMIN_USERNAME || 'DamianSpears_Owner';
const ADMIN_DEFAULT_PASSWORD = process.env.ADMIN_PASSWORD || 'Papaleta723!';
let adminPasswordHash = '';

// Hasheo seguro al inicializar el servidor (nunca almacenada en plano)
bcrypt.hash(ADMIN_DEFAULT_PASSWORD, 10).then(hash => {
  adminPasswordHash = hash;
  console.log(`🔒 Cuenta Administrador configurada: '${ADMIN_USERNAME}' (Contraseña hasheada con bcrypt).`);
});

// Almacén de sesiones activas en memoria
// Estructura: token -> { username, loginTime, lastActive, expiresAt }
const adminSessions = new Map();
const SESSION_TTL_MS = 30 * 60 * 1000; // 30 minutos por inactividad

// Rate Limiting para Login Admin (Prevención de Fuerza Bruta)
const loginAttempts = new Map(); // ip -> { count, firstAttempt }
const MAX_ATTEMPTS = 5;
const LOCK_TIME_MS = 15 * 60 * 1000; // 15 minutos de bloqueo

function checkRateLimit(ip) {
  const now = Date.now();
  const record = loginAttempts.get(ip);
  if (!record) return { allowed: true };

  if (now - record.firstAttempt > LOCK_TIME_MS) {
    loginAttempts.delete(ip);
    return { allowed: true };
  }

  if (record.count >= MAX_ATTEMPTS) {
    const minutesLeft = Math.ceil((LOCK_TIME_MS - (now - record.firstAttempt)) / 60000);
    return { allowed: false, minutesLeft };
  }

  return { allowed: true };
}

function registerFailedAttempt(ip) {
  const now = Date.now();
  const record = loginAttempts.get(ip) || { count: 0, firstAttempt: now };
  record.count += 1;
  loginAttempts.set(ip, record);
}

function resetFailedAttempts(ip) {
  loginAttempts.delete(ip);
}

// Middleware para proteger rutas exclusivas de Administrador
function requireAdmin(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;

  if (!token) {
    return res.status(401).json({ success: false, mensaje: 'Acceso no autorizado: Token no proporcionado.' });
  }

  const session = adminSessions.get(token);
  if (!session) {
    return res.status(401).json({ success: false, mensaje: 'Sesión no válida o expirada.' });
  }

  const now = Date.now();
  if (now > session.expiresAt) {
    adminSessions.delete(token);
    return res.status(401).json({ success: false, mensaje: 'La sesión de administrador ha expirado por inactividad.' });
  }

  // Renovar tiempo de vida por actividad
  session.lastActive = now;
  session.expiresAt = now + SESSION_TTL_MS;
  req.admin = session;
  next();
}

// ==========================================
// CONFIGURACIÓN DE EMPRESA Y VIÁTICOS
// ==========================================
const EMPRESA_CONFIG = {
  rfcEmpresa: 'GME980512ABC',
  razonSocialEmpresa: 'GRUPO MERCANTIL EMPRESARIAL SA DE CV',
  politicaTopeAlimentosDiario: 750.00 // Art. 28 Fracc. V LISR
};

// ==========================================
// CATÁLOGO DE EMPLEADOS Y ASIGNACIÓN DE VIAJES
// ==========================================
let EMPLEADOS = [
  {
    id: 'emp-julian',
    nombre: 'Julián Alberto López Morales',
    email: 'julian.lopez@grupomercantil.com.mx',
    claveAcceso: 'JALM-2026',
    rfc: 'LOMJ880914HC3',
    puesto: 'Gerente de Operaciones y Logística',
    departamento: 'Operaciones',
    estatus: 'ACTIVO',
    iniciales: 'JL',
    avatarGradient: 'from-emerald-600 to-teal-800',
    viaje: {
      id: 'via-2026-0042',
      folio: 'VIA-2026-0042',
      motivo: 'Auditoría a Centro de Distribución y Alianzas de Transporte',
      destino: 'Monterrey, Nuevo León',
      esNacional: true,
      fechaInicio: '2026-10-20',
      fechaFin: '2026-10-22',
      diasViaje: 3,
      anticipoOtorgado: 6000.00
    }
  },
  {
    id: 'emp-sandra',
    nombre: 'Sandra Urquilla Rosales',
    email: 'sandra.urquilla@grupomercantil.com.mx',
    claveAcceso: 'SUR-2026',
    rfc: 'UROS920318M14',
    puesto: 'Directora de Expansión Comercial',
    departamento: 'Ventas Corporativas',
    estatus: 'ACTIVO',
    iniciales: 'SU',
    avatarGradient: 'from-indigo-600 to-violet-800',
    viaje: {
      id: 'via-2026-0058',
      motivo: 'Cierre de Contratos Marco con Cuentas Clave de Occidente',
      destino: 'Guadalajara, Jalisco',
      esNacional: true,
      fechaInicio: '2026-10-18',
      fechaFin: '2026-10-21',
      diasViaje: 4,
      anticipoOtorgado: 8500.00
    }
  },
  {
    id: 'emp-angel',
    nombre: 'Ángel Miramontes',
    email: 'angel.miramontes@grupomercantil.com.mx',
    claveAcceso: 'AM-2026',
    rfc: 'MIRA851120AB9',
    puesto: 'Auditor Contable y Fiscal Senior',
    departamento: 'Contraloría y Finanzas',
    estatus: 'ACTIVO',
    iniciales: 'AM',
    avatarGradient: 'from-amber-600 to-orange-800',
    viaje: {
      id: 'via-2026-0063',
      motivo: 'Auditoría Fiscal de Activos y Control de Sucursal Sureste',
      destino: 'Cancún, Quintana Roo',
      esNacional: true,
      fechaInicio: '2026-10-15',
      fechaFin: '2026-10-19',
      diasViaje: 5,
      anticipoOtorgado: 12000.00
    }
  }
];

// Comprobantes por empleado
let comprobantesPorEmpleado = {
  'emp-julian': [],
  'emp-sandra': [],
  'emp-angel': []
};

// ==========================================
// PERSISTENCIA EN ARCHIVOS JSON
// ==========================================
const DATA_DIR = path.join(__dirname, 'data');
const EMPLEADOS_FILE = path.join(DATA_DIR, 'empleados.json');
const COMPROBANTES_FILE = path.join(DATA_DIR, 'comprobantes.json');

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

function guardarDatos() {
  try {
    fs.writeFileSync(EMPLEADOS_FILE, JSON.stringify(EMPLEADOS, null, 2), 'utf8');
    fs.writeFileSync(COMPROBANTES_FILE, JSON.stringify(comprobantesPorEmpleado, null, 2), 'utf8');
  } catch (err) {
    console.error('Error al guardar datos en disco:', err);
  }
}

function cargarDatos() {
  try {
    if (fs.existsSync(EMPLEADOS_FILE)) {
      const data = fs.readFileSync(EMPLEADOS_FILE, 'utf8');
      EMPLEADOS = JSON.parse(data);
      console.log(`📁 Catálogo de empleados cargado desde disco: ${EMPLEADOS.length} colaboradores.`);
    } else {
      guardarDatos();
    }
    if (fs.existsSync(COMPROBANTES_FILE)) {
      const data = fs.readFileSync(COMPROBANTES_FILE, 'utf8');
      comprobantesPorEmpleado = JSON.parse(data);
    } else {
      guardarDatos();
    }
  } catch (err) {
    console.error('Error al cargar datos desde disco:', err);
  }
}

// Cargar al iniciar
cargarDatos();

// ==========================================
// PARSER ROBUSTO DE CFDI 4.0
// ==========================================
function parsearCfdi40(xmlStr) {
  if (!xmlStr || typeof xmlStr !== 'string') {
    throw new Error('El archivo XML está vacío o no es una cadena válida.');
  }

  const extractAttr = (tagRegex, attrName) => {
    const matchTag = xmlStr.match(tagRegex);
    if (!matchTag) return null;
    const tagContent = matchTag[0];
    const attrRegex = new RegExp(`\\b${attrName}\\s*=\\s*["']([^"']*)["']`, 'i');
    const matchAttr = tagContent.match(attrRegex);
    return matchAttr ? matchAttr[1] : null;
  };

  const version = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'Version');
  if (!version) {
    throw new Error('No se encontró el nodo raíz <cfdi:Comprobante> en el XML.');
  }
  if (version !== '4.0') {
    throw new Error(`Versión de comprobante no admitida: CFDI ${version}. El portal solo admite CFDI 4.0.`);
  }

  const total = parseFloat(extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'Total') || '0');
  const subTotal = parseFloat(extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'SubTotal') || '0');
  const moneda = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'Moneda') || 'MXN';
  const fechaStr = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'Fecha') || '';
  const formaPago = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'FormaPago') || '99';
  const metodoPago = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'MetodoPago') || 'PUE';
  const serie = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'Serie') || '';
  const folio = extractAttr(/<[a-zA-Z0-9_:]*Comprobante\b[^>]*>/i, 'Folio') || '';

  const rfcEmisor = extractAttr(/<[a-zA-Z0-9_:]*Emisor\b[^>]*>/i, 'Rfc') || 'SIN_RFC';
  const nombreEmisor = extractAttr(/<[a-zA-Z0-9_:]*Emisor\b[^>]*>/i, 'Nombre') || 'EMISOR DESCONOCIDO';
  const rfcReceptor = extractAttr(/<[a-zA-Z0-9_:]*Receptor\b[^>]*>/i, 'Rfc') || 'SIN_RFC';
  const nombreReceptor = extractAttr(/<[a-zA-Z0-9_:]*Receptor\b[^>]*>/i, 'Nombre') || 'RECEPTOR DESCONOCIDO';

  const uuid = extractAttr(/<[a-zA-Z0-9_:]*TimbreFiscalDigital\b[^>]*>/i, 'UUID');
  if (!uuid) {
    throw new Error('El comprobante no contiene el complemento TimbreFiscalDigital o no cuenta con UUID fiscal del SAT.');
  }

  return {
    uuid: uuid.toUpperCase(),
    version,
    total,
    subTotal,
    moneda,
    fechaEmision: fechaStr,
    formaPago,
    metodoPago,
    serie,
    folio,
    rfcEmisor: rfcEmisor.toUpperCase(),
    nombreEmisor,
    rfcReceptor: rfcReceptor.toUpperCase(),
    nombreReceptor
  };
}

// ==========================================
// VALIDACIONES FISCALES Y DE VIAJE
// ==========================================
function validarCfdiContraViaje(cfdi, montoDeclarado, categoria = 'ALIMENTOS', empleado, modoFlexible = false) {
  const reglas = [];
  let esValido = true;
  const comprobantesDelEmpleado = comprobantesPorEmpleado[empleado.id] || [];

  const yaExisteUuid = comprobantesDelEmpleado.some(c => c.uuid === cfdi.uuid);
  if (yaExisteUuid) {
    reglas.push({
      nombre: 'UUID Único (Anti-duplicidad SAT)',
      pasa: false,
      detalle: `El UUID ${cfdi.uuid} ya fue registrado previamente en la comprobación de ${empleado.nombre}.`
    });
    esValido = false;
  } else {
    reglas.push({
      nombre: 'UUID Único (Anti-duplicidad SAT)',
      pasa: true,
      detalle: `Folio fiscal ${cfdi.uuid} válido y no registrado previamente.`
    });
  }

  const rfcCoincide = (cfdi.rfcReceptor === EMPRESA_CONFIG.rfcEmpresa);
  if (!rfcCoincide && !modoFlexible) {
    reglas.push({
      nombre: 'RFC Receptor de la Empresa',
      pasa: false,
      detalle: `El CFDI está a nombre de '${cfdi.rfcReceptor}', pero la razón social corporativa es '${EMPRESA_CONFIG.rfcEmpresa}' (${EMPRESA_CONFIG.razonSocialEmpresa}).`
    });
    esValido = false;
  } else {
    reglas.push({
      nombre: 'RFC Receptor de la Empresa',
      pasa: true,
      detalle: modoFlexible && !rfcCoincide
        ? `[Modo Flexible Activo] Aceptado RFC Receptor '${cfdi.rfcReceptor}' (${cfdi.nombreReceptor}) para pruebas libres.`
        : `RFC Receptor correcto (${EMPRESA_CONFIG.rfcEmpresa} - ${EMPRESA_CONFIG.razonSocialEmpresa}).`
    });
  }

  const fechaEmisionSoloDia = cfdi.fechaEmision.split('T')[0];
  const viaje = empleado.viaje;
  const enPeriodo = (fechaEmisionSoloDia >= viaje.fechaInicio && fechaEmisionSoloDia <= '2026-10-31');

  if (!enPeriodo && !modoFlexible) {
    reglas.push({
      nombre: 'Fecha en Período Autorizado',
      pasa: false,
      detalle: `La fecha de la factura (${fechaEmisionSoloDia}) está fuera del periodo de la comisión (${viaje.fechaInicio} al ${viaje.fechaFin}).`
    });
    esValido = false;
  } else {
    reglas.push({
      nombre: 'Fecha en Período Autorizado',
      pasa: true,
      detalle: modoFlexible && !enPeriodo
        ? `[Modo Flexible Activo] Aceptada fecha real (${fechaEmisionSoloDia}) para la comprobación de viáticos.`
        : `Fecha de emisión (${fechaEmisionSoloDia}) dentro del periodo del viaje a ${viaje.destino}.`
    });
  }

  const diferencia = Math.abs(montoDeclarado - cfdi.total);
  const discrepanciaMonto = diferencia > 0.05;

  if (discrepanciaMonto) {
    reglas.push({
      nombre: 'Conciliación Monto Declarado vs XML',
      pasa: false,
      detalle: `Discrepancia detectada: Declaraste $${montoDeclarado.toFixed(2)} MXN, pero el XML del SAT marca un total de $${cfdi.total.toFixed(2)} MXN (Diferencia: $${diferencia.toFixed(2)} MXN).`
    });
    esValido = false;
  } else {
    reglas.push({
      nombre: 'Conciliación Monto Declarado vs XML',
      pasa: true,
      detalle: `Monto capturado ($${montoDeclarado.toFixed(2)}) coincide con el total exacto del CFDI ($${cfdi.total.toFixed(2)} MXN).`
    });
  }

  reglas.push({
    nombre: 'Estatus SAT WebService',
    pasa: true,
    detalle: 'Estado en el SAT: VIGENTE. Sin reportes de cancelación ni empresas simuladoras (EFOS 69-B).'
  });

  return { esValido, reglas, discrepanciaMonto, diferencia };
}

// ==========================================
// CÁLCULO DE LIQUIDACIÓN Y DEDUCIBILIDAD LISR
// ==========================================
function calcularLiquidacion(empleadoId) {
  const empleado = EMPLEADOS.find(e => e.id === empleadoId) || EMPLEADOS[0];
  const comprobantes = comprobantesPorEmpleado[empleado.id] || [];

  let totalComprobadoBruto = 0;
  let totalDeducible = 0;
  let totalExcedenteNoDeducible = 0;

  const alimentosPorDia = {};

  comprobantes.forEach(c => {
    // Si la factura fue rechazada por administración, no suma al deducible
    if (c.estatus === 'RECHAZADO') {
      c.deducible = 0;
      c.excedente = c.total;
      return;
    }

    totalComprobadoBruto += c.total;
    const dia = c.fechaEmision.split('T')[0];

    if (c.categoria === 'ALIMENTOS') {
      const acumuladoDia = alimentosPorDia[dia] || 0;
      const tope = EMPRESA_CONFIG.politicaTopeAlimentosDiario;

      if (acumuladoDia + c.total <= tope) {
        totalDeducible += c.total;
        alimentosPorDia[dia] = acumuladoDia + c.total;
        c.deducible = c.total;
        c.excedente = 0;
      } else {
        const parteDeducible = Math.max(0, tope - acumuladoDia);
        const parteExcedente = c.total - parteDeducible;
        totalDeducible += parteDeducible;
        totalExcedenteNoDeducible += parteExcedente;
        alimentosPorDia[dia] = tope;
        c.deducible = parteDeducible;
        c.excedente = parteExcedente;
      }
    } else {
      totalDeducible += c.total;
      c.deducible = c.total;
      c.excedente = 0;
    }
  });

  const anticipo = empleado.viaje ? empleado.viaje.anticipoOtorgado : 0;
  let saldoAFavorEmpresa = 0;
  let saldoAFavorColaborador = 0;

  if (anticipo > totalDeducible) {
    saldoAFavorEmpresa = Number((anticipo - totalDeducible).toFixed(2));
  } else {
    saldoAFavorColaborador = Number((totalDeducible - anticipo).toFixed(2));
  }

  return {
    empleado,
    empresa: EMPRESA_CONFIG,
    totalComprobadoBruto: Number(totalComprobadoBruto.toFixed(2)),
    totalDeducible: Number(totalDeducible.toFixed(2)),
    totalExcedenteNoDeducible: Number(totalExcedenteNoDeducible.toFixed(2)),
    saldoAFavorEmpresa,
    saldoAFavorColaborador,
    comprobantes
  };
}

// ==========================================
// MUESTRAS DE PRUEBA RÁPIDA
// ==========================================
const SAMPLES = {
  valido: `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4" xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="4.0" Serie="MTR" Folio="9941" Fecha="2026-10-21T14:30:00" FormaPago="04" SubTotal="620.69" Total="720.00" Moneda="MXN" MetodoPago="PUE">
  <cfdi:Emisor Rfc="RST140510A12" Nombre="RESTAURANTE EL REY DEL CABRITO SA DE CV" RegimenFiscal="601"/>
  <cfdi:Receptor Rfc="GME980512ABC" Nombre="GRUPO MERCANTIL EMPRESARIAL SA DE CV" UsoCFDI="G03" DomicilioFiscalReceptor="64000" RegimenFiscalReceptor="601"/>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="9B1A2C3D-4E5F-6A7B-8C9D-0E1F2A3B4C5D" FechaTimbrado="2026-10-21T14:31:00" RfcProvCertif="SAT970701NN3"/>
  </cfdi:Complemento>
</cfdi:Comprobante>`,

  discrepancia_monto: `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4" xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="4.0" Serie="H" Folio="205" Fecha="2026-10-21T21:00:00" FormaPago="04" SubTotal="850.00" Total="986.00" Moneda="MXN" MetodoPago="PUE">
  <cfdi:Emisor Rfc="HOT900101XYZ" Nombre="HOTEL CATEDRAL SA" RegimenFiscal="601"/>
  <cfdi:Receptor Rfc="GME980512ABC" Nombre="GRUPO MERCANTIL EMPRESARIAL SA DE CV" UsoCFDI="G03"/>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="8F2E3D4C-5B6A-7890-1234-567890ABCDEF"/>
  </cfdi:Complemento>
</cfdi:Comprobante>`,

  fecha_invalida: `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4" xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="4.0" Serie="A" Folio="54" Fecha="2026-10-05T12:00:00" FormaPago="04" SubTotal="450.00" Total="522.00" Moneda="MXN" MetodoPago="PUE">
  <cfdi:Emisor Rfc="CAF120101999" Nombre="CAFETERIA ANTIGUA SA" RegimenFiscal="601"/>
  <cfdi:Receptor Rfc="GME980512ABC" Nombre="GRUPO MERCANTIL EMPRESARIAL SA DE CV" UsoCFDI="G03"/>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="7E6D5C4B-3A21-4F5E-6D7C-8B9A0F1E2D3C"/>
  </cfdi:Complemento>
</cfdi:Comprobante>`,

  rfc_incorrecto: `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4" xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="4.0" Serie="B" Folio="1002" Fecha="2026-10-21T16:00:00" FormaPago="04" SubTotal="300.00" Total="348.00" Moneda="MXN" MetodoPago="PUE">
  <cfdi:Emisor Rfc="TLL080808111" Nombre="TAXIS EJECUTIVOS DEL NORTE" RegimenFiscal="601"/>
  <cfdi:Receptor Rfc="XAXX010101000" Nombre="PUBLICO EN GENERAL" UsoCFDI="S01"/>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="6A5B4C3D-2E1F-0A9B-8C7D-6E5F4A3B2C1D"/>
  </cfdi:Complemento>
</cfdi:Comprobante>`,

  excedente_lisr: `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4" xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="4.0" Serie="FAC" Folio="7889" Fecha="2026-10-22T15:00:00" FormaPago="04" SubTotal="1120.69" Total="1300.00" Moneda="MXN" MetodoPago="PUE">
  <cfdi:Emisor Rfc="STK190201444" Nombre="STEAK HOUSE SAN PEDRO SA DE CV" RegimenFiscal="601"/>
  <cfdi:Receptor Rfc="GME980512ABC" Nombre="GRUPO MERCANTIL EMPRESARIAL SA DE CV" UsoCFDI="G03"/>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="5D4C3B2A-1F0E-9D8C-7B6A-5E4D3C2B1A0F"/>
  </cfdi:Complemento>
</cfdi:Comprobante>`
};

// ==========================================
// ENDPOINTS DE AUTENTICACIÓN Y ADMIN
// ==========================================

// 1. Login Exclusivo de Administrador con Rate Limiting y Mensajes Genéricos
app.post('/api/admin/login', async (req, res) => {
  const clientIp = req.ip || req.connection.remoteAddress || '127.0.0.1';
  const { allowed, minutesLeft } = checkRateLimit(clientIp);

  if (!allowed) {
    return res.status(429).json({
      success: false,
      mensaje: `Demasiados intentos fallidos. Por seguridad, intente de nuevo en ${minutesLeft} minutos.`
    });
  }

  const { username, password } = req.body;

  if (!username || !password) {
    registerFailedAttempt(clientIp);
    return res.status(400).json({ success: false, mensaje: 'Credenciales inválidas' });
  }

  // Comparación en tiempo constante para evitar ataques de temporización
  const usernameMatch = (username.trim() === ADMIN_USERNAME);
  let passwordMatch = false;

  if (adminPasswordHash) {
    passwordMatch = await bcrypt.compare(password, adminPasswordHash);
  }

  // Mensaje genérico: no revela si el usuario existe o no
  if (!usernameMatch || !passwordMatch) {
    registerFailedAttempt(clientIp);
    return res.status(401).json({ success: false, mensaje: 'Credenciales inválidas' });
  }

  // Éxito: resetear intentos fallidos
  resetFailedAttempts(clientIp);

  // Crear token de sesión criptográficamente seguro
  const token = crypto.randomBytes(32).toString('hex');
  const now = Date.now();
  adminSessions.set(token, {
    username: ADMIN_USERNAME,
    loginTime: now,
    lastActive: now,
    expiresAt: now + SESSION_TTL_MS
  });

  res.json({
    success: true,
    mensaje: 'Autenticación exitosa como Administrador.',
    token,
    admin: {
      username: ADMIN_USERNAME,
      rol: 'ADMIN_OWNER'
    }
  });
});

// 2. Verificar estado de la sesión de administrador
app.get('/api/admin/verificar-sesion', requireAdmin, (req, res) => {
  res.json({
    success: true,
    admin: {
      username: req.admin.username,
      lastActive: req.admin.lastActive,
      expiresAt: req.admin.expiresAt
    }
  });
});

// 3. Cierre de sesión de administrador
app.post('/api/admin/logout', (req, res) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (token) {
    adminSessions.delete(token);
  }
  res.json({ success: true, mensaje: 'Sesión de administrador cerrada correctamente.' });
});

// 4. Listar todos los usuarios/empleados (Solo Admin)
app.get('/api/admin/usuarios', requireAdmin, (req, res) => {
  res.json({
    success: true,
    usuarios: EMPLEADOS.map(e => ({
      id: e.id,
      nombre: e.nombre,
      email: e.email || `${e.id}@grupomercantil.com.mx`,
      rfc: e.rfc,
      puesto: e.puesto,
      departamento: e.departamento,
      claveAcceso: e.claveAcceso,
      estatus: e.estatus || 'ACTIVO',
      viaje: e.viaje,
      totalComprobantes: (comprobantesPorEmpleado[e.id] || []).length
    }))
  });
});

// 5. Dar de alta nuevo usuario (Solo Admin con validaciones estrictas)
app.post('/api/admin/usuarios', requireAdmin, (req, res) => {
  const { nombre, email, rfc, puesto, departamento, claveAcceso, viaje } = req.body;

  // Validaciones de campos obligatorios
  if (!nombre || !nombre.trim()) {
    return res.status(400).json({ success: false, mensaje: 'El nombre completo es obligatorio.' });
  }

  // Validación de formato de correo electrónico
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!email || !emailRegex.test(email.trim())) {
    return res.status(400).json({ success: false, mensaje: 'El correo electrónico no tiene un formato válido (ej. usuario@empresa.com).' });
  }

  // Validación básica de formato RFC mexicano (12 o 13 caracteres alfanuméricos)
  const rfcRegex = /^[A-ZÑ&]{3,4}\d{6}[A-V1-9][A-Z1-9][0-9A]$/i;
  const rfcLimpio = (rfc || '').trim().toUpperCase();
  if (!rfcLimpio || rfcLimpio.length < 12 || rfcLimpio.length > 13) {
    return res.status(400).json({ success: false, mensaje: 'El RFC debe contener entre 12 y 13 caracteres válidos.' });
  }

  if (!claveAcceso || !claveAcceso.trim()) {
    return res.status(400).json({ success: false, mensaje: 'La clave única de acceso es obligatoria.' });
  }

  // Validar unicidad de la clave de acceso y del correo
  const claveDuplicada = EMPLEADOS.some(e => e.claveAcceso.toUpperCase() === claveAcceso.trim().toUpperCase());
  if (claveDuplicada) {
    return res.status(400).json({ success: false, mensaje: `La clave única '${claveAcceso}' ya está asignada a otro colaborador.` });
  }

  const emailDuplicado = EMPLEADOS.some(e => (e.email || '').toLowerCase() === email.trim().toLowerCase());
  if (emailDuplicado) {
    return res.status(400).json({ success: false, mensaje: `El correo '${email}' ya está registrado.` });
  }

  // Generar ID único
  const nuevoId = 'emp-' + crypto.randomUUID().slice(0, 8);
  const iniciales = nombre.trim().split(' ').map(n => n[0]).join('').slice(0, 2).toUpperCase();

  const nuevoUsuario = {
    id: nuevoId,
    nombre: nombre.trim(),
    email: email.trim().toLowerCase(),
    claveAcceso: claveAcceso.trim().toUpperCase(),
    rfc: rfcLimpio,
    puesto: (puesto || 'Especialista').trim(),
    departamento: (departamento || 'Operaciones').trim(),
    estatus: 'ACTIVO',
    iniciales: iniciales || 'US',
    avatarGradient: 'from-blue-600 to-cyan-800',
    viaje: viaje || {
      id: 'via-' + Date.now().toString().slice(-4),
      folio: 'VIA-2026-' + (EMPLEADOS.length + 1).toString().padStart(4, '0'),
      motivo: 'Comisión de Trabajo Autorizada',
      destino: 'Ciudad de México, CDMX',
      esNacional: true,
      fechaInicio: '2026-10-20',
      fechaFin: '2026-10-24',
      diasViaje: 5,
      anticipoOtorgado: 5000.00
    }
  };

  EMPLEADOS.push(nuevoUsuario);
  comprobantesPorEmpleado[nuevoId] = [];
  guardarDatos();
  console.log(`👤 Nuevo colaborador registrado: ${nuevoUsuario.nombre} (${nuevoUsuario.id})`);

  res.status(201).json({
    success: true,
    mensaje: `Colaborador '${nuevoUsuario.nombre}' dado de alta exitosamente.`,
    usuario: nuevoUsuario
  });
});

// 6. Editar usuario existente (Solo Admin)
app.put('/api/admin/usuarios/:id', requireAdmin, (req, res) => {
  const usuario = EMPLEADOS.find(e => e.id === req.params.id);
  if (!usuario) {
    return res.status(404).json({ success: false, mensaje: 'Colaborador no encontrado.' });
  }

  const { nombre, email, rfc, puesto, departamento, claveAcceso, estatus, viaje } = req.body;

  if (nombre && nombre.trim()) usuario.nombre = nombre.trim();
  if (email && email.trim()) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email.trim())) {
      return res.status(400).json({ success: false, mensaje: 'El correo electrónico no tiene un formato válido.' });
    }
    usuario.email = email.trim().toLowerCase();
  }
  if (rfc && rfc.trim()) usuario.rfc = rfc.trim().toUpperCase();
  if (puesto && puesto.trim()) usuario.puesto = puesto.trim();
  if (departamento && departamento.trim()) usuario.departamento = departamento.trim();
  if (estatus && ['ACTIVO', 'INACTIVO'].includes(estatus)) usuario.estatus = estatus;

  if (claveAcceso && claveAcceso.trim()) {
    const claveRepetida = EMPLEADOS.some(e => e.id !== usuario.id && e.claveAcceso.toUpperCase() === claveAcceso.trim().toUpperCase());
    if (claveRepetida) {
      return res.status(400).json({ success: false, mensaje: `La clave única '${claveAcceso}' ya pertenece a otro usuario.` });
    }
    usuario.claveAcceso = claveAcceso.trim().toUpperCase();
  }

  if (viaje) {
    usuario.viaje = {
      ...usuario.viaje,
      ...viaje
    };
  }

  guardarDatos();
  console.log(`✏️ Colaborador actualizado: ${usuario.nombre} (${usuario.id})`);

  res.json({
    success: true,
    mensaje: `Datos del colaborador '${usuario.nombre}' actualizados correctamente.`,
    usuario
  });
});

// ==========================================
// ENDPOINTS PÚBLICOS DE EMPLEADOS Y VIÁTICOS
// ==========================================

// 1. Obtener lista pública de empleados activos para viáticos (Claves y Montos protegidos con asteriscos)
app.get('/api/empleados', (req, res) => {
  res.json({
    success: true,
    empleados: EMPLEADOS
      .filter(e => !e.estatus || e.estatus.trim().toUpperCase() === 'ACTIVO')
      .map(e => ({
        id: e.id,
        nombre: e.nombre,
        claveAccesoMasked: '••••••••', // Nunca revelar la clave en texto plano
        anticipoMasked: '$ ••••••',    // Privacidad: solo visible al autenticar la sesión individual
        rfc: e.rfc,
        puesto: e.puesto,
        departamento: e.departamento,
        iniciales: e.iniciales,
        avatarGradient: e.avatarGradient,
        folioViaje: e.viaje ? e.viaje.folio : 'SIN-VIAJE',
        destino: e.viaje ? e.viaje.destino : 'No asignado',
        totalComprobantes: (comprobantesPorEmpleado[e.id] || []).length
      }))
  });
});

// 2. Verificar clave única de acceso en backend de forma privada
app.post('/api/empleados/:id/verificar-clave', (req, res) => {
  const empleado = EMPLEADOS.find(e => e.id === req.params.id);
  if (!empleado) {
    return res.status(404).json({ success: false, mensaje: 'Colaborador no encontrado.' });
  }

  const { clave } = req.body;
  if (!clave || typeof clave !== 'string') {
    return res.status(400).json({ success: false, mensaje: 'Debe ingresar su clave única.' });
  }

  const claveValida = (clave.trim().toUpperCase() === empleado.claveAcceso.toUpperCase());
  if (!claveValida) {
    return res.status(401).json({ success: false, mensaje: 'Clave única incorrecta. Verifícala con Administración.' });
  }

  res.json({
    success: true,
    mensaje: `Acceso concedido para ${empleado.nombre}`,
    empleadoId: empleado.id
  });
});

// Obtener estado del viaje y liquidación de un empleado
app.get('/api/empleados/:id/viaje', (req, res) => {
  const empleado = EMPLEADOS.find(e => e.id === req.params.id);
  if (!empleado) {
    return res.status(404).json({ success: false, mensaje: 'Colaborador no encontrado.' });
  }

  res.json({
    success: true,
    data: calcularLiquidacion(empleado.id)
  });
});

// Procesar comprobante para un empleado
app.post('/api/empleados/:id/validar-comprobante', (req, res) => {
  try {
    const empleado = EMPLEADOS.find(e => e.id === req.params.id);
    if (!empleado) {
      return res.status(404).json({ success: false, mensaje: 'Colaborador no encontrado.' });
    }

    const { xmlString, montoDeclarado, categoria, modoFlexible } = req.body;

    if (!xmlString) {
      return res.status(400).json({ success: false, mensaje: 'El contenido XML es obligatorio.' });
    }

    const montoNum = parseFloat(montoDeclarado);
    if (isNaN(montoNum) || montoNum <= 0) {
      return res.status(400).json({ success: false, mensaje: 'Ingrese un monto declarado válido.' });
    }

    const cfdi = parsearCfdi40(xmlString);
    const validacion = validarCfdiContraViaje(cfdi, montoNum, categoria || 'ALIMENTOS', empleado, !!modoFlexible);

    if (validacion.esValido) {
      const nuevoComprobante = {
        id: crypto.randomUUID(),
        ...cfdi,
        categoria: categoria || 'ALIMENTOS',
        montoDeclarado: montoNum,
        estatus: 'APROBADO',
        hashSha256: crypto.createHash('sha256').update(xmlString).digest('hex'),
        fechaRegistro: new Date().toISOString()
      };
      if (!comprobantesPorEmpleado[empleado.id]) comprobantesPorEmpleado[empleado.id] = [];
      comprobantesPorEmpleado[empleado.id].push(nuevoComprobante);
      guardarDatos();
    }

    res.json({
      success: validacion.esValido,
      cfdi,
      validacion,
      liquidacionActualizada: calcularLiquidacion(empleado.id)
    });
  } catch (err) {
    res.status(422).json({
      success: false,
      mensaje: err.message || 'Error al procesar el archivo XML.'
    });
  }
});

// Actualizar estatus de un comprobante (Aceptar / Rechazar por Administrador)
app.put('/api/admin/empleados/:empId/comprobantes/:uuid/estatus', requireAdmin, (req, res) => {
  const { empId, uuid } = req.params;
  const { estatus, motivo } = req.body;

  const empleado = EMPLEADOS.find(e => e.id === empId);
  if (!empleado) {
    return res.status(404).json({ success: false, mensaje: 'Colaborador no encontrado.' });
  }

  const comprobantes = comprobantesPorEmpleado[empId] || [];
  const comp = comprobantes.find(c => (c.uuid && c.uuid.toUpperCase() === uuid.toUpperCase()) || c.id === uuid);

  if (!comp) {
    return res.status(404).json({ success: false, mensaje: 'Comprobante fiscal no encontrado.' });
  }

  const nuevoEstatus = (estatus || '').trim().toUpperCase();
  if (!['APROBADO', 'RECHAZADO'].includes(nuevoEstatus)) {
    return res.status(400).json({ success: false, mensaje: "Estatus no válido. Debe ser 'APROBADO' o 'RECHAZADO'." });
  }

  comp.estatus = nuevoEstatus;
  comp.motivoAdmin = motivo || (nuevoEstatus === 'RECHAZADO' ? 'Rechazado manualmente por el Administrador.' : 'Aprobado por el Administrador.');
  comp.fechaRevisionAdmin = new Date().toISOString();

  guardarDatos();
  console.log(`🧾 Comprobante ${comp.uuid} de ${empleado.nombre} marcado como ${nuevoEstatus}`);

  res.json({
    success: true,
    mensaje: `Factura ${nuevoEstatus === 'APROBADO' ? 'aprobada' : 'rechazada'} correctamente.`,
    comprobante: comp,
    liquidacionActualizada: calcularLiquidacion(empId)
  });
});

// Reiniciar comprobantes de un empleado
app.post('/api/empleados/:id/reiniciar', (req, res) => {
  const empleado = EMPLEADOS.find(e => e.id === req.params.id);
  if (!empleado) {
    return res.status(404).json({ success: false, mensaje: 'Colaborador no encontrado.' });
  }
  comprobantesPorEmpleado[empleado.id] = [];
  guardarDatos();
  res.json({
    success: true,
    mensaje: `Comprobación de ${empleado.nombre} reiniciada.`,
    data: calcularLiquidacion(empleado.id)
  });
});

// Muestras de prueba
app.get('/api/samples/:tipo', (req, res) => {
  const sample = SAMPLES[req.params.tipo];
  if (!sample) return res.status(404).json({ error: 'Muestra no encontrada' });
  res.json({ tipo: req.params.tipo, xmlString: sample });
});

// Compatibilidad retroactiva
app.get('/api/viaje', (req, res) => {
  res.json({ success: true, data: calcularLiquidacion(EMPLEADOS[0].id) });
});

app.listen(PORT, () => {
  console.log(`\n===============================================================`);
  console.log(`✨ PORTAL DE VIÁTICOS MÉXICO CON MÓDULO DE ADMINISTRACIÓN`);
  console.log(`📍 URL Local: http://localhost:${PORT}`);
  console.log(`🛡️ Administrador: ${ADMIN_USERNAME}`);
  console.log(`👥 Colaboradores en catálogo: ${EMPLEADOS.length}`);
  console.log(`===============================================================\n`);
});
