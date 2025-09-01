import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion, NubeInterSostenimiento } from '../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelSostenimientoImportService {

  constructor() {}

  async importOperacionesFromExcel(file: File): Promise<any[]> {
    const data: any[] = [];

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

    const ejecutadoSheet = workbook.Sheets['EJECUTADOSOS'];
    const estadosSheet = workbook.Sheets['ESTADOSSOS'];
    const checklistSheet = workbook.Sheets['CHECK LISTSOS'];

    const ejecutadoData = XLSX.utils.sheet_to_json<any>(ejecutadoSheet);
    const estadosData = XLSX.utils.sheet_to_json<any>(estadosSheet);
    const checklistData = XLSX.utils.sheet_to_json<any>(checklistSheet);

    // Mapear operaciones básicas
    ejecutadoData.forEach(row => {
      const idOp = row['ID Operación'];
      if (!data.find(op => op.__excel_id === idOp)) {
        data.push({
          __excel_id: idOp,
          operacion: {
            turno: row['Turno'] || '',
            equipo: row['Equipo'] || '',
            codigo: row['Código'] || '',
            empresa: row['Empresa'] || '',
            fecha: row['Fecha'] || '',
            tipo_operacion: row['Tipo Operación'] || '',
            estado: row['Estado'] || '',
            envio: 0
          },
          horometros: [],
          checklists: [],
          estados: []
        });
      }
    });

    // Mapear horómetros
    ejecutadoData.forEach(row => {
      const op = data.find(op => op.__excel_id === row['ID Operación']);
      if (op) {
        Object.keys(row).forEach(key => {
          const match = key.match(/^Horómetro (.+) - Inicial$/);
          if (match) {
            const nombre = match[1].replace(/_/g, ' ');
            op.horometros.push({
              nombre,
              inicial: Number(row[`Horómetro ${nombre} - Inicial`] || 0),
              final: Number(row[`Horómetro ${nombre} - Final`] || 0),
              EstaOP: row[`Horómetro ${nombre} - Operativo`] === 'Sí' ? 1 : 0,
              EstaINOP: row[`Horómetro ${nombre} - Operativo`] === 'No' ? 1 : 0
            });
          }
        });
      }
    });

    // Mapear checklist
    checklistData.forEach(row => {
      const op = data.find(op => op.__excel_id === row['ID Operación']);
      if (op && !row['Mensaje']) {
        op.checklists.push({
          descripcion: row['Descripción'] || '',
          decision: row['Decisión'] === 'Aprobado' ? 1 : 0,
          observacion: row['Observación'] || '',
          categoria: row['Categoría'] || ''
        });
      }
    });

    // Mapear estados y sostenimientos
    estadosData.forEach(row => {
      const op = data.find(op => op.__excel_id === row['ID Operación']);
      if (!op || row['Mensaje']) return;

      // Buscar o crear estado
      let estado = op.estados.find((e: any) => e.numero === Number(row['Número Estado'] || 0));
      if (!estado) {
        estado = {
          numero: Number(row['Número Estado'] || 0),
          estado: row['Estado'] || '',
          codigo: row['Código Estado'] || '',
          hora_inicio: row['Hora Inicio'] || '',
          hora_final: row['Hora Final'] || '',
          sostenimientos: []
        };
        op.estados.push(estado);
      }

      const zona = (row['Sost. - Zona'] || '').toString().trim();
      const tipoLabor = (row['Sost. - Tipo Labor'] || '').toString().trim();

      let sost = estado.sostenimientos.find((s: any) =>
        s.zona?.toString().trim() === zona && s.tipo_labor?.toString().trim() === tipoLabor
      );

      if (!sost && zona) {
        sost = {
          zona,
          tipo_labor: tipoLabor,
          labor: row['Sost. - Labor'] || '',
          veta: row['Sost. - Veta'] || '',
          nivel: row['Sost. - Nivel'] || '',
          tipo_perforacion: row['Sost. - Tipo Perforación'] || '',
          observacion: row['Sost. - Observación'] || '',
          inter_sostenimientos: []
        };
        estado.sostenimientos.push(sost);
      }

      if (sost) {
        const inter: Omit<NubeInterSostenimiento, 'id' | 'sostenimiento_id'> = {
          codigo_actividad: row['Ejecutado - Código Actividad'] || '',
          nivel: row['Ejecutado - Nivel'] || '',
          labor: row['Ejecutado - Labor'] || '',
          seccion_de_labor: row['Ejecutado - Sección'] || '',
          nbroca: Number(row['Ejecutado - N° Broca'] || 0),
          ntaladro: Number(row['Ejecutado - N° Taladro'] || 0),
          material: row['Ejecutado - Material'] || '',
          longitud_perforacion: Number(row['Ejecutado - Longitud'] || 0),
          malla_instalada: row['Ejecutado - Malla Instalada'] || '',
          metros_perforados: Number(row['Ejecutado - Metros perforados'] || 0),
          detalles_trabajo_realizado: row['Ejecutado - Detalles'] || ''
        };
        sost.inter_sostenimientos.push(inter);
      }
    });

    // Limpiar y devolver en formato API
    return data.map(op => {
      const { __excel_id, ...rest } = op;
      return rest;
    });
  }
}
