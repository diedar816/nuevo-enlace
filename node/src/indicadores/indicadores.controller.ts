
import { Controller, Get, Query } from '@nestjs/common';
import * as path from 'path';
import * as Excel from 'exceljs';

// Normaliza valores para convertirlos a número o null
function numberOrNull(x: any): number | null {
  if (x === null || x === undefined) return null;
  if (typeof x === 'number' && Number.isFinite(x)) return x;
  const s = String(x).trim();
  if (!s) return null;
  const norm = s.replace(/\s/g, '').replace(/\./g, '').replace(/,/g, '.');
  const n = Number(norm);
  return Number.isFinite(n) ? n : null;
}

@Controller('indicadores')
export class IndicadoresController {
  // Salud
  @Get('health')
  health() {
    return { ok: true, at: new Date().toISOString() };
  }

  // Lee el Excel real y devuelve las 2 columnas necesarias
  @Get('tabla-resultado')
  async getTablaResultado(@Query('sheet') sheetName?: string) {
    // OJO: el Processor lo generó con ESPACIO y extensión .xlsx (minúscula)
    const FILE_PATH = path.join(
      'C:',
      'Proyecto TICs',
      'Nuevo Enlace',
      'Tabla_ Resultado.xlsx',
    );

    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(FILE_PATH);

    // Usa hoja indicada (?sheet=...), o intenta 'Todos', o la primera
    const ws =
      (sheetName && wb.getWorksheet(sheetName)) ||
      wb.getWorksheet('Todos') ||
      wb.worksheets[0];

    if (!ws) return [];

    // Encabezados en fila 1
    const headers: string[] = [];
    ws.getRow(1).eachCell((cell, i) => {
      headers[i - 1] = String(cell.value ?? '').trim();
    });

    // Filas a objetos por nombre de encabezado
    const rows: Record<string, any>[] = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const obj: Record<string, any> = {};
      headers.forEach((h, idx) => (obj[h] = ws.getRow(r).getCell(idx + 1).value));
      rows.push(obj);
    }

    // Solo lo que usa el front, en número
    const out = rows.map((r) => ({
      TOTAL_DIAS_VENTANILLA: numberOrNull(r['TOTAL_DIAS_VENTANILLA']),
      TIEMPO_EN_MESA_DE_CREACION: numberOrNull(r['TIEMPO_EN_MESA_DE_CREACION']),
    }));

    return out;
  }
}
