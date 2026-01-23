
import { Injectable } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import * as Excel from 'exceljs';

/** ---------------------------
 *  Paths absolutos (Windows)
 * --------------------------*/
const BASE_DIR = path.join('C:', 'Proyecto TICs', 'Nuevo Enlace');
const ENTRADA  = path.join(BASE_DIR, 'Tabla Modelo.XLSX');
const SALIDA   = path.join(BASE_DIR, 'Tabla_ Resultado.xlsx');

/** ---------------------------
 *  Utilidades de texto/número
 * --------------------------*/
function removeAccentsUpper(x: any): string {
  if (x === null || x === undefined) return '';
  return String(x).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

const LONG_RAD = 14;

function scientificToIntegerString(s: string): string {
  const m = s.trim().toLowerCase().match(/^([+-]?)(\d+)(?:\.(\d+))?e([+-]?\d+)$/);
  if (!m) return s;
  const sign = m[1] === '-' ? '-' : '';
  const intPart = m[2] ?? '0';
  const frac = m[3] ?? '';
  const exp = parseInt(m[4], 10);

  let digits = intPart + frac;
  const dotPos = intPart.length;
  let newDotPos = dotPos + exp;

  if (exp >= 0) {
    if (newDotPos >= digits.length) {
      digits = digits + '0'.repeat(newDotPos - digits.length);
      return (sign ? '-' : '') + digits;
    }
    const whole = digits.slice(0, newDotPos);
    return (sign ? '-' : '') + (whole || '0');
  } else {
    newDotPos = dotPos + exp;
    if (newDotPos <= 0) return '0';
    const whole = digits.slice(0, newDotPos);
    return (sign ? '-' : '') + (whole || '0');
  }
}

function normalizarRadicado(x: any): string | null {
  if (x === null || x === undefined) return null;
  let s = String(x).trim();
  try {
    if (/e/i.test(s)) {
      s = scientificToIntegerString(s);
    } else {
      const sNum = s.replace(/[,\s]/g, '');
      if (/^-?\d+(\.\d+)?$/.test(sNum)) {
        const n = Math.trunc(Number(sNum));
        s = String(n);
      }
    }
  } catch {}
  const digits = s.replace(/\D/g, '');
  return digits ? digits.padStart(LONG_RAD, '0') : s;
}

/** ---------------------------
 *  Fechas y Festivos Colombia
 * --------------------------*/

/** Normaliza valor excel a Date (sin hora). Soporta: Date, serial Excel, yyyy-mm-dd, dd/mm/yyyy, dd-mm-yyyy */
function parseDateFromExcel(val: any): Date | null {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  if (typeof val === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
    const ms = Math.round(val * 86400000);
    const d = new Date(excelEpoch.getTime() + ms);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  if (typeof val === 'string') {
    const s = val.trim();

    // yyyy-mm-dd
    let mm = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (mm) {
      const [_, Y, M, D] = mm;
      return new Date(parseInt(Y), parseInt(M) - 1, parseInt(D));
    }
    // dd/mm/yyyy o dd-mm-yyyy
    mm = s.match(/^(\d{1,2})\D(\d{1,2})\D(\d{4})$/);
    if (mm) {
      const [_, d, mth, Y] = mm;
      return new Date(parseInt(Y), parseInt(mth) - 1, parseInt(d));
    }
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    }
  }
  return null;
}

function dateToKey(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function addDays(d: Date, days: number): Date {
  const r = new Date(d);
  r.setDate(r.getDate() + days);
  return new Date(r.getFullYear(), r.getMonth(), r.getDate());
}
function diasNaturales(a: Date | null, b: Date | null, inclusive = false): number | null {
  if (!a || !b) return null;
  const diff = Math.round((b.getTime() - a.getTime()) / 86400000);
  return inclusive ? diff + 1 : diff;
}
function isSaturday(d: Date): boolean { return d.getDay() === 6; }
function isSunday(d: Date): boolean { return d.getDay() === 0; }

/** Pascua (Domingo) - algoritmo gregoriano (Meeus/Jones/Butcher) */
function easterSunday(year: number): Date {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31); // 3=Mar, 4=Abr
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

/** Ley Emiliani: mover al lunes siguiente */
function moveToNextMonday(d: Date): Date {
  const wd = d.getDay(); // 0=Dom,1=Lun,...6=Sab
  if (wd === 1) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  let r = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  while (r.getDay() !== 1) r = addDays(r, 1);
  return r;
}

function fixedHoliday(year: number, month: number, day: number): Date {
  return new Date(year, month - 1, day);
}

/** Festivos Colombia por año (fijos, Semana Santa, y movidos por Emiliani) */
function getColombiaHolidaysForYear(year: number): Date[] {
  const easter = easterSunday(year);
  const juevesSanto = addDays(easter, -3);
  const viernesSanto = addDays(easter, -2);

  const fixed = [
    fixedHoliday(year, 1, 1),   // Año Nuevo
    fixedHoliday(year, 5, 1),   // Trabajo
    fixedHoliday(year, 7, 20),  // Independencia
    fixedHoliday(year, 8, 7),   // Boyacá
    fixedHoliday(year, 12, 8),  // Inmaculada
    fixedHoliday(year, 12, 25), // Navidad
  ];

  const emilianiBase = [
    fixedHoliday(year, 1, 6),   // Reyes
    fixedHoliday(year, 3, 19),  // San José
    fixedHoliday(year, 6, 29),  // San Pedro y San Pablo
    fixedHoliday(year, 8, 15),  // Asunción
    fixedHoliday(year, 10, 12), // Raza
    fixedHoliday(year, 11, 1),  // Todos los Santos
    fixedHoliday(year, 11, 11), // Cartagena
  ];
  const emiliani = emilianiBase.map(moveToNextMonday);

  const ascension = moveToNextMonday(addDays(easter, 39)); // jueves → lunes
  const corpus = moveToNextMonday(addDays(easter, 60));    // jueves → lunes
  const sagradoCorazon = moveToNextMonday(addDays(easter, 68)); // domingo → lunes

  return [...fixed, juevesSanto, viernesSanto, ascension, corpus, sagradoCorazon, ...emiliani];
}

/** Build base de festivos para un rango de años */
function buildHolidaySetForRange(minYear: number, maxYear: number): Set<string> {
  const set = new Set<string>();
  for (let y = minYear; y <= maxYear; y++) {
    for (const d of getColombiaHolidaysForYear(y)) set.add(dateToKey(d));
  }
  return set;
}

/** Estado global de festivos (se recalcula según rango detectado) */
let HOLIDAYS_SET_RANGE: Set<string> | null = null;
let HOLIDAYS_RANGE: { min: number; max: number } | null = null;

function ensureHolidayRange(minYear: number, maxYear: number) {
  if (!HOLIDAYS_SET_RANGE || !HOLIDAYS_RANGE || minYear < HOLIDAYS_RANGE.min || maxYear > HOLIDAYS_RANGE.max) {
    const buildMin = HOLIDAYS_RANGE ? Math.min(HOLIDAYS_RANGE.min, minYear) : minYear;
    const buildMax = HOLIDAYS_RANGE ? Math.max(HOLIDAYS_RANGE.max, maxYear) : maxYear;
    HOLIDAYS_SET_RANGE = buildHolidaySetForRange(buildMin, buildMax);
    HOLIDAYS_RANGE = { min: buildMin, max: buildMax };
  }
}

function isHolidayColombia(d: Date): boolean {
  if (!HOLIDAYS_SET_RANGE) return false;
  return HOLIDAYS_SET_RANGE.has(dateToKey(d));
}

/** ===========================================
 * INYECCIÓN DE FESTIVOS (JSON y Hoja Excel)
 * ===========================================*/
type HolidayAction = 'ADD' | 'REMOVE';
type HolidayInjected = { add: Set<string>; remove: Set<string> };

function parseKeyOrNull(s: any): string | null {
  if (!s) return null;
  const str = String(s).trim();
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? str : null;
}

/** Cargar inyección desde JSON */
function loadHolidaysFromJsonFile(filePath: string): HolidayInjected | null {
  try {
    const full = path.isAbsolute(filePath) ? filePath : path.join(process.cwd(), filePath);
    if (!fs.existsSync(full)) {
      console.warn('Archivo JSON de festivos no encontrado:', full);
      return null;
    }
    const raw = fs.readFileSync(full, 'utf8');
    const data = JSON.parse(raw);

    const out: HolidayInjected = { add: new Set<string>(), remove: new Set<string>() };

    // Soportar formatos:
    // 1) ["YYYY-MM-DD", ...]
    // 2) [{ "date":"YYYY-MM-DD", "action":"ADD"|"REMOVE" }, ...]
    // 3) { "holidays": [ ... ] }
    const list: any[] = Array.isArray(data) ? data
      : Array.isArray(data?.holidays) ? data.holidays
      : [];

    for (const item of list) {
      if (typeof item === 'string') {
        const k = parseKeyOrNull(item);
        if (k) out.add.add(k);
      } else if (item && typeof item === 'object') {
        const k = parseKeyOrNull(item.date ?? item.fecha);
        const action = String(item.action ?? item.accion ?? 'ADD').toUpperCase();
        if (!k) continue;
        if (action === 'REMOVE') out.remove.add(k);
        else out.add.add(k);
      }
    }
    return out;
  } catch (e) {
    console.warn('No se pudo cargar JSON de festivos:', e);
    return null;
  }
}

/** Cargar inyección desde hoja "Festivos" del mismo Excel de entrada
 *  Espera encabezados (insensible a mayúsculas):
 *    - fecha   (YYYY-MM-DD)
 *    - accion  (ADD|REMOVE) opcional; por defecto ADD
 */
function loadHolidaysFromFestivosSheet(wb: Excel.Workbook): HolidayInjected | null {
  const ws = wb.getWorksheet('Festivos');
  if (!ws) return null;

  // Detectar columnas por nombre en fila 1
  const headers: Record<string, number> = {};
  ws.getRow(1).eachCell((cell, col) => {
    const name = String(cell.value ?? '').trim().toLowerCase();
    if (name) headers[name] = col;
  });

  // Columnas esperadas
  const colFecha = headers['fecha'] ?? headers['date'];
  const colAccion = headers['accion'] ?? headers['action'];

  if (!colFecha) {
    console.warn('Hoja "Festivos": no se encontró columna "fecha". Se ignora la hoja.');
    return null;
  }

  const out: HolidayInjected = { add: new Set<string>(), remove: new Set<string>() };

  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    let key: string | null = null;

    // La fecha puede venir como Date, serial o string
    const rawFecha = row.getCell(colFecha).value;
    const d = parseDateFromExcel(rawFecha);
    if (d) key = dateToKey(d);
    else {
      const s = parseKeyOrNull(rawFecha);
      if (s) key = s;
    }
    if (!key) continue;

    const actRaw = colAccion ? row.getCell(colAccion).value : null;
    const act = String(actRaw ?? 'ADD').toUpperCase();

    if (act === 'REMOVE') out.remove.add(key);
    else out.add.add(key);
  }

  return out;
}

/** Aplica inyección a la base (según modo) */
type HolidaysMode = 'merge' | 'override';
function applyInjectedHolidays(mode: HolidaysMode, injected: HolidayInjected | null) {
  if (!injected) return;
  if (!HOLIDAYS_SET_RANGE) HOLIDAYS_SET_RANGE = new Set<string>();

  if (mode === 'override') {
    HOLIDAYS_SET_RANGE = new Set<string>(); // arranca desde vacío
    for (const k of injected.add) HOLIDAYS_SET_RANGE.add(k);
    // REMOVE no tiene sentido en override, pero lo respetamos por si acaso
    for (const k of injected.remove) HOLIDAYS_SET_RANGE.delete(k);
  } else {
    // merge: parte de la base automática y aplica cambios
    for (const k of injected.add) HOLIDAYS_SET_RANGE.add(k);
    for (const k of injected.remove) HOLIDAYS_SET_RANGE.delete(k);
  }
}

/** ---------------------------
 *  Hábiles Colombia (excluye sábados, domingos y festivos)
 * --------------------------*/
const INCLUDE_END = true;
function diasHabilesColombia(a: Date | null, b: Date | null, includeEnd = INCLUDE_END): number | null {
  if (!a || !b) return null;

  let start = a;
  let end = b;

  if (includeEnd) {
    if (end >= start) end = addDays(end, 1);
    else start = addDays(start, 1);
  }

  let sign = 1;
  if (end < start) {
    sign = -1;
    const t = start; start = end; end = t;
  }

  let cnt = 0;
  for (let d = start; d < end; d = addDays(d, 1)) {
    if (!isSaturday(d) && !isSunday(d) && !isHolidayColombia(d)) cnt++;
  }
  return sign * cnt;
}

/** ---------------------------
 *  Detección de columnas y helpers
 * --------------------------*/
function detectarColumna(headers: string[], candidatos: string[]): string | null {
  const mapa = new Map<string, string>();
  headers.forEach(h => mapa.set(h.toLowerCase().trim(), h));
  for (const c of candidatos) {
    const k = c.toLowerCase().trim();
    if (mapa.has(k)) return mapa.get(k)!;
  }
  for (const c of candidatos) {
    const k = c.toLowerCase().trim();
    for (const [norm, real] of mapa.entries()) {
      if (norm.includes(k)) return real;
    }
  }
  return null;
}

function dedupeKeepLastByName<T extends Record<string, any>>(rows: T[], _target: string): T[] {
  return rows;
}

/** ---------------------------
 *  CLI args simples
 * --------------------------*/
function getArg(name: string): string | undefined {
  // soporta --name=value o --name value
  const argv = process.argv.slice(2);
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith(name + '=')) return a.split('=').slice(1).join('=');
    if (a === name) return argv[i + 1];
  }
  return undefined;
}

/** ---------------------------
 *  Servicio principal
 * --------------------------*/
@Injectable()
export class ProcessorService {
  async run(): Promise<void> {
    console.log('== Inicio del script ==');
    console.log(`Carpeta base: ${BASE_DIR}`);
    console.log(`Archivo entrada: ${ENTRADA}`);

    // CLI: inyectar festivos
    const holidaysPath = getArg('--holidays') ?? getArg('-H'); // ruta JSON opcional
    const holidaysMode = ((getArg('--holidaysMode') ?? 'merge').toLowerCase()) as HolidaysMode; // merge|override
    if (!['merge','override'].includes(holidaysMode)) {
      throw new Error(`holidaysMode inválido: ${holidaysMode}. Use "merge" u "override".`);
    }

    // Chequeo de acceso
    try {
      fs.accessSync(ENTRADA, fs.constants.R_OK);
      const st = fs.statSync(ENTRADA);
      console.log('Input OK | tamaño:', Math.round(st.size / 1024), 'KB');
    } catch (e) {
      console.error('No se puede leer el archivo de entrada:', e);
      throw e;
    }

    // Lectura por streaming
    const wbIn = new Excel.Workbook();
    console.time('excel:stream-read');
    await wbIn.xlsx.read(fs.createReadStream(ENTRADA));
    console.timeEnd('excel:stream-read');

    const wsIn = wbIn.worksheets[0];
    if (!wsIn) throw new Error('No se encontró hoja en el archivo de entrada.');

    // Encabezados
    const headers: string[] = [];
    wsIn.getRow(1).eachCell({ includeEmpty: false }, (cell, colNumber) => {
      headers[colNumber - 1] = String(cell.value ?? '').trim();
    });
    const headerCols = headers
      .map((name, idx) => ({ name, col: idx + 1 }))
      .filter(h => h.name && h.name.length > 0);

    const totalRows: number = (wsIn as any).actualRowCount ?? wsIn.rowCount;

    // Filas como objetos
    console.time('build:rows');
    const rows: Record<string, any>[] = [];
    for (let r = 2; r <= totalRows; r++) {
      const xlRow = wsIn.getRow(r);
      const obj: Record<string, any> = {};
      for (const hc of headerCols) obj[hc.name] = xlRow.getCell(hc.col).value;
      rows.push(obj);
    }
    console.timeEnd('build:rows');
    console.log(`Filas: ${rows.length} | Columnas: ${headerCols.length}`);

    // Detección de columnas
    console.time('detect:columns');
    const headerNames = headerCols.map(h => h.name);
    const col_radicado  = detectarColumna(headerNames, ['numero_radicado','nro_radicado','radicado']);
    const col_dep       = detectarColumna(headerNames, ['nombre_dependencia','dependencia']);
    const col_asigna    = detectarColumna(headerNames, ['fecha_asignacion','fecha_asigna','f_asignacion','asignacion']);
    const col_caso_crea = detectarColumna(headerNames, ['CASO_FECHA_CREACION','caso_fecha_creacion','fecha_creacion_caso','f_caso_crea']);
    console.timeEnd('detect:columns');

    console.log('Columnas detectadas:');
    console.log('  numero_radicado     =>', col_radicado);
    console.log('  nombre_dependencia  =>', col_dep);
    console.log('  fecha_asignacion    =>', col_asigna);
    console.log('  CASO_FECHA_CREACION =>', col_caso_crea);

    const faltantes = [
      ['numero_radicado', col_radicado],
      ['nombre_dependencia', col_dep],
      ['fecha_asignacion', col_asigna],
      ['CASO_FECHA_CREACION', col_caso_crea],
    ].filter(([_, v]) => !v).map(([k]) => k);
    if (faltantes.length) throw new Error(`Faltan columnas imprescindibles: ${JSON.stringify(faltantes)}`);

    // Normalizaciones y captura de rango de años
    console.time('normalize:dates+radicado');
    let minYear = Number.POSITIVE_INFINITY;
    let maxYear = Number.NEGATIVE_INFINITY;
    for (const row of rows) {
      const dAsig = parseDateFromExcel(row[col_asigna!]);
      const dCrea = parseDateFromExcel(row[col_caso_crea!]);
      row['_fecha_asignacion'] = dAsig;
      row['_caso_fecha_creacion'] = dCrea;

      if (dAsig) { minYear = Math.min(minYear, dAsig.getFullYear()); maxYear = Math.max(maxYear, dAsig.getFullYear()); }
      if (dCrea) { minYear = Math.min(minYear, dCrea.getFullYear()); maxYear = Math.max(maxYear, dCrea.getFullYear()); }

      const radNorm = normalizarRadicado(row[col_radicado!]);
      row['_rad_str'] = radNorm ?? '';
      row[col_radicado!] = row['_rad_str'];
    }
    if (!isFinite(minYear) || !isFinite(maxYear)) {
      const y = new Date().getFullYear(); minYear = y; maxYear = y;
    }
    // Construir base automática de festivos según rango detectado
    ensureHolidayRange(minYear, maxYear);

    // Inyección desde hoja Excel (si existe)
    const injectedExcel = loadHolidaysFromFestivosSheet(wbIn);

    // Inyección desde JSON (si se pasó por CLI)
    const injectedJson = holidaysPath ? loadHolidaysFromJsonFile(holidaysPath) : null;

    // Merge ambas inyecciones en un único paquete
    const mergedInjected: HolidayInjected | null = ((): HolidayInjected | null => {
      if (!injectedExcel && !injectedJson) return null;
      const add = new Set<string>();
      const remove = new Set<string>();
      if (injectedExcel) {
        for (const k of injectedExcel.add) add.add(k);
        for (const k of injectedExcel.remove) remove.add(k);
      }
      if (injectedJson) {
        for (const k of injectedJson.add) add.add(k);
        for (const k of injectedJson.remove) remove.add(k);
      }
      return { add, remove };
    })();

    // Aplicar inyección según modo (merge|override)
    applyInjectedHolidays(holidaysMode, mergedInjected);

    console.timeEnd('normalize:dates+radicado');

    // Máscaras por dependencia
    const depNormArr = rows.map(row => removeAccentsUpper(row[col_dep!]));
    const maskVent = depNormArr.map(s => /VENTANILL/.test(s));
    const maskMesa = depNormArr.map(s => /\bMESA\b/.test(s));

    // Índice MESA por radicado
    console.time('build:mesa-index');
    const mesaAsig = new Map<string, Date[]>();
    rows.forEach((row, i) => {
      if (!maskMesa[i]) return;
      const rad = String(row[col_radicado!]);
      const d = row['_fecha_asignacion'] as Date | null;
      if (!d) return;
      if (!mesaAsig.has(rad)) mesaAsig.set(rad, []);
      mesaAsig.get(rad)!.push(d);
    });
    function mesaRefDate(radicado: string, ventDate: Date): Date | null {
      const fechas = mesaAsig.get(String(radicado)) ?? [];
      const valids = fechas.filter(Boolean);
      if (!valids.length) return null;
      const posteriores = valids.filter(d => d >= ventDate);
      if (posteriores.length) return posteriores.reduce((a, b) => (a < b ? a : b));
      return valids.reduce((a, b) => (a < b ? a : b));
    }
    console.timeEnd('build:mesa-index');

    // TOTAL_DIAS_VENTANILLA
    console.time('calc:vent-mesa');
    let pareosLogrados = 0;
    rows.forEach((row, i) => {
      row['TOTAL_DIAS_VENTANILLA_NATURALES'] = null;
      row['TOTAL_DIAS_VENTANILLA_HABILES'] = null;

      if (maskVent[i] && row['_fecha_asignacion']) {
        const rad = String(row[col_radicado!]);
        const ventDate = row['_fecha_asignacion'] as Date;
        const mesaDate = mesaRefDate(rad, ventDate);
        if (mesaDate) {
          pareosLogrados++;
          row['TOTAL_DIAS_VENTANILLA_NATURALES'] = diasNaturales(ventDate, mesaDate, false);
          row['TOTAL_DIAS_VENTANILLA_HABILES']   = diasHabilesColombia(ventDate, mesaDate, true);
        }
      }
      row['TOTAL_DIAS_VENTANILLA'] = row['TOTAL_DIAS_VENTANILLA_HABILES'];
    });
    console.timeEnd('calc:vent-mesa');
    console.log(`Pareos VENT ↔ MESA logrados: ${pareosLogrados}/${maskVent.filter(Boolean).length}`);

    // TIEMPO_EN_MESA_DE_CREACION
    console.time('calc:mesa-horiz');
    rows.forEach((row, i) => {
      row['TIEMPO_EN_MESA_DE_CREACION_NATURALES'] = null;
      row['TIEMPO_EN_MESA_DE_CREACION_HABILES'] = null;

      if (maskMesa[i]) {
        const a = row['_fecha_asignacion'] as Date | null;
        const b = row['_caso_fecha_creacion'] as Date | null;
        row['TIEMPO_EN_MESA_DE_CREACION_NATURALES'] = diasNaturales(a, b, false);
        row['TIEMPO_EN_MESA_DE_CREACION_HABILES']   = diasHabilesColombia(a, b, true);
      }
      row['TIEMPO_EN_MESA_DE_CREACION'] = row['TIEMPO_EN_MESA_DE_CREACION_HABILES'];
    });
    console.timeEnd('calc:mesa-horiz');

    // ALERTAS
    console.time('alerts');
    rows.forEach((row, i) => {
      const avisos: string[] = [];
      if (maskVent[i]) {
        const tdvH = row['TOTAL_DIAS_VENTANILLA_HABILES'];
        if (tdvH == null && row['_fecha_asignacion']) {
          avisos.push('VENT: Sin MESA para pareo (mismo radicado)');
        }
        if (typeof tdvH === 'number' && tdvH < 0) {
          avisos.push('VENT: Diferencia negativa (MESA < VENT)');
        }
        const cantMesa = (mesaAsig.get(String(row[col_radicado!])) ?? []).filter(Boolean).length;
        if (cantMesa > 1 && tdvH != null) {
          avisos.push(`VENT: MESA múltiple (${cantMesa}) → se eligió la más cercana ≥ VENT`);
        }
      }
      if (maskMesa[i]) {
        if (!row['_fecha_asignacion'] || !row['_caso_fecha_creacion']) {
          avisos.push('MESA: Fecha(s) faltante(s)');
        } else {
          const tHab = row['TIEMPO_EN_MESA_DE_CREACION_HABILES'];
          if (typeof tHab === 'number' && tHab < 0) {
            avisos.push('MESA: Diferencia negativa (CREACIÓN < ASIGNACIÓN)');
          }
        }
      }
      row['ALERTA'] = avisos.length ? avisos.join('; ') : '';
    });
    console.timeEnd('alerts');

    // Filas pertinentes
    const pertinentes = rows.map(r =>
      Boolean(r['ALERTA']) ||
      r['TOTAL_DIAS_VENTANILLA'] != null ||
      r['TIEMPO_EN_MESA_DE_CREACION'] != null
    );
    const df_pert = rows.filter((_, i) => pertinentes[i]);

    // Resumen por radicado
    console.time('build:summary');
    type TmpRow = {
      [k: string]: any;
      tipo: 'VENTANILLA' | 'MESA';
      TOTAL_DIAS_VENTANILLA?: number | null;
      TIEMPO_EN_MESA_DE_CREACION?: number | null;
    };
    const tmp: TmpRow[] = rows
      .filter((_, i) => maskVent[i] || maskMesa[i])
      .map(r => {
        const tipo = removeAccentsUpper(r[col_dep!]).startsWith('VENTANILL') ? 'VENTANILLA' : 'MESA';
        return {
          [col_radicado!]: r[col_radicado!],
          [col_dep!]: r[col_dep!],
          TOTAL_DIAS_VENTANILLA: r['TOTAL_DIAS_VENTANILLA'] ?? null,
          TIEMPO_EN_MESA_DE_CREACION: r['TIEMPO_EN_MESA_DE_CREACION'] ?? null,
          tipo,
        };
      });

    const summaryMap = new Map<string, Record<string, any>>();
    for (const r of tmp) {
      const rad = String(r[col_radicado!]);
      if (!summaryMap.has(rad)) summaryMap.set(rad, { [col_radicado!]: rad });
      const agg = summaryMap.get(rad)!;
      const sfx = r.tipo;
      const k1 = `TOTAL_DIAS_VENTANILLA_${sfx}`;
      const k2 = `TIEMPO_EN_MESA_DE_CREACION_${sfx}`;
      const a = r.TOTAL_DIAS_VENTANILLA;
      const b = r.TIEMPO_EN_MESA_DE_CREACION;
      agg[k1] = (typeof agg[k1] === 'number' ? agg[k1] : Number.NEGATIVE_INFINITY);
      agg[k2] = (typeof agg[k2] === 'number' ? agg[k2] : Number.NEGATIVE_INFINITY);
      if (typeof a === 'number') agg[k1] = Math.max(agg[k1], a);
      if (typeof b === 'number') agg[k2] = Math.max(agg[k2], b);
    }
    const summary = Array.from(summaryMap.values()).map(o => {
      for (const k of Object.keys(o)) if (o[k] === Number.NEGATIVE_INFINITY) o[k] = null;
      return o;
    });
    console.timeEnd('build:summary');

    // Auditoría de columnas detectadas
    const detalles_cols = [
      { campo_logico: 'numero_radicado', columna_en_archivo: col_radicado },
      { campo_logico: 'nombre_dependencia', columna_en_archivo: col_dep },
      { campo_logico: 'fecha_asignacion', columna_en_archivo: col_asigna },
      { campo_logico: 'CASO_FECHA_CREACION', columna_en_archivo: col_caso_crea },
      { campo_logico: 'Regla TOTAL_DIAS_VENTANILLA', columna_en_archivo: 'MESA.fecha_asignacion - VENT.fecha_asignacion (mismo radicado)' },
      { campo_logico: 'Regla TIEMPO_EN_MESA_DE_CREACION', columna_en_archivo: 'MESA.CASO_FECHA_CREACION - MESA.fecha_asignacion' },
    ];

    // Auditoría de pareo
    console.time('build:auditoria-pareo');
    const audRows: any[] = [];
    rows.forEach((row, i) => {
      if (!maskVent[i]) return;
      const rad = row[col_radicado!];
      const ventDate = row['_fecha_asignacion'] as Date | null;
      const mesas = (mesaAsig.get(String(rad)) ?? []).filter(Boolean);
      const mesaSel = ventDate ? mesaRefDate(String(rad), ventDate) : null;

      let estado = 'SIN_PAREO';
      let motivo = 'Sin fechas MESA';

      if (mesas.length) {
        if (!ventDate) {
          estado = 'SIN_PAREO';
          motivo = 'VENT sin fecha';
        } else if (mesaSel) {
          if (mesaSel >= ventDate) {
            estado = 'PAREO';
            motivo = 'MESA ≥ VENT';
          } else {
            estado = 'PAREO';
            motivo = 'MESA mínima < VENT';
          }
        } else {
          estado = 'SIN_PAREO';
          motivo = 'No se encontró MESA ≥ VENT';
        }
      }

      const fechasDisp = mesas.map(d => dateToKey(d)).sort().join(', ');

      audRows.push({
        [col_radicado!]: rad,
        VENT_FECHA_ASIGNACION: ventDate ?? null,
        MESA_FECHAS_DISPONIBLES: fechasDisp,
        MESA_FECHA_SELECCIONADA: mesaSel ?? null,
        ESTADO_PAREO: estado,
        MOTIVO: motivo,
      });
    });
    const auditoria_pareo = audRows;
    console.timeEnd('build:auditoria-pareo');

    /** ---------------------------
     *  Escritura Excel salida
     * --------------------------*/
    console.log('Guardando Excel de salida...');
    const wbOut = new Excel.Workbook();

    function addSheetFromObjects(
      wb: Excel.Workbook,
      name: string,
      data: Record<string, any>[],
    ): Excel.Worksheet {
      const ws = wb.addWorksheet(name);
      if (!data.length) return ws;
      const cols = Object.keys(data[0]);
      ws.addRow(cols);
      for (const r of data) ws.addRow(cols.map(c => r[c]));
      return ws;
    }

    // Ocultar técnicas
    const dropCols = new Set<string>([
      'TOTAL_DIAS_VENTANILLA_HABILES',
      'TOTAL_DIAS_VENTANILLA_NATURALES',
      'TIEMPO_EN_MESA_DE_CREACION_HABILES',
      'TIEMPO_EN_MESA_DE_CREACION_NATURALES',
      'N',
    ]);
    function cleanRows(arr: Record<string, any>[]): Record<string, any>[] {
      return arr.map(orig => {
        const o: Record<string, any> = {};
        for (const k of Object.keys(orig)) {
          if (k.startsWith('_')) continue;
          if (dropCols.has(k)) continue;
          o[k] = orig[k];
        }
        return o;
      });
    }

    let df_out = cleanRows(rows);
    let df_pert_out = cleanRows(df_pert);

    // Reordenar visibles
    const tailCols = ['ALERTA', 'TIEMPO_EN_MESA_DE_CREACION', 'TOTAL_DIAS_VENTANILLA'];
    function reorderRows(arr: Record<string, any>[]): Record<string, any>[] {
      if (!arr.length) return arr;
      const all = Object.keys(arr[0]);
      const head = all.filter(c => !tailCols.includes(c));
      const cols = [...head, ...tailCols.filter(c => all.includes(c))];
      return arr.map(r => {
        const o: Record<string, any> = {};
        for (const c of cols) o[c] = r[c];
        return o;
      });
    }
    df_out = reorderRows(df_out);
    df_pert_out = reorderRows(df_pert_out);

    const wsTodos   = addSheetFromObjects(wbOut, 'Todos', df_out);
    const wsRev     = addSheetFromObjects(wbOut, 'Registros_para_revisar', df_pert_out);
    const wsResumen = addSheetFromObjects(wbOut, 'Resumen_rad_duplicados', summary);
    const wsCols    = addSheetFromObjects(wbOut, 'Columnas_detectadas', detalles_cols);
    const wsAud     = addSheetFromObjects(wbOut, 'Auditoria_pareo', auditoria_pareo);

    // Formato de fechas y alineación
    const fmtFecha = 'yyyy-mm-dd';
    function headerIndexMap(ws: Excel.Worksheet, nombres: Set<string>): Map<string, number> {
      const map = new Map<string, number>();
      const header = ws.getRow(1);
      header.eachCell((cell, col) => {
        const val = String(cell.value ?? '').trim();
        if (nombres.has(val)) map.set(val, col);
      });
      return map;
    }
    function formatearFechas(ws: Excel.Worksheet, columnas: string[]) {
      if (!ws || ws.rowCount < 2) return;
      const idx = headerIndexMap(ws, new Set(columnas));
      idx.forEach((colIndex) => {
        for (let r = 2; r <= ws.rowCount; r++) {
          const cell = ws.getRow(r).getCell(colIndex);
          if (typeof cell.value === 'string') {
            const d = parseDateFromExcel(cell.value);
            if (d) cell.value = d;
          }
          cell.numFmt = fmtFecha;
        }
      });
    }
    function alinearHoja(ws: Excel.Worksheet) {
      const align = { horizontal: 'distributed', vertical: 'top', wrapText: true } as Excel.Alignment;
      ws.eachRow(row => row.eachCell(cell => (cell.alignment = align)));
    }

    const fechaColsTodosRev = [col_asigna!, col_caso_crea!];
    formatearFechas(wsTodos, fechaColsTodosRev);  alinearHoja(wsTodos);
    formatearFechas(wsRev,   fechaColsTodosRev);  alinearHoja(wsRev);
    alinearHoja(wsResumen);
    alinearHoja(wsCols);
    formatearFechas(wsAud, ['VENT_FECHA_ASIGNACION', 'MESA_FECHA_SELECCIONADA']); alinearHoja(wsAud);

    // Forzar numero_radicado como texto
    function forzarRadicadoTexto(ws: Excel.Worksheet, colName: string) {
      const header = headerIndexMap(ws, new Set([colName]));
      const colIndex = header.get(colName);
      if (!colIndex) return;
      for (let r = 2; r <= ws.rowCount; r++) {
        const cell = ws.getRow(r).getCell(colIndex);
        if (cell.value !== null && cell.value !== undefined) {
          cell.value = `'` + String(cell.value);
          cell.numFmt = '@';
        }
      }
    }
    for (const ws of [wsTodos, wsResumen, wsAud]) {
      forzarRadicadoTexto(ws, col_radicado!);
    }

    console.time('write:file');
    await wbOut.xlsx.writeFile(SALIDA);
    console.timeEnd('write:file');
    console.log('Archivo generado:', SALIDA);
  }
}
