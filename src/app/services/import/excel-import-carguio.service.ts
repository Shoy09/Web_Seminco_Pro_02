import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion, NubeCarguio } from '../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelCarguioImportService {

  constructor() {}

  async importOperacionesFromExcel(file: File): Promise<any[]> {
    const data: any[] = [];

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

    const ejecutadoSheet = workbook.Sheets['EJECUTADOFR'];
    const estadosSheet = workbook.Sheets['ESTADOSFR'];
    const checklistSheet = workbook.Sheets['CHECK LISTFR'];

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

    // Mapear estados y carguios
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
          carguios: []
        };
        op.estados.push(estado);
      }

      // Crear registro de carguio
      const carg: Omit<NubeCarguio, 'id' | 'estado_id'> = {
        tipo_labor: row['Carguío - Tipo Labor'] || '',
        labor: row['Carguío - Labor'] || '',
        tipo_labor_manual: row['Carguío - Tipo Labor Manual'] || '',
        labor_manual: row['Carguío - Labor Manual'] || '',
        ncucharas: Number(row['Carguío - Nº Cucharas'] || 0),
        observacion: row['Carguío - Observación'] || ''
      };

      estado.carguios.push(carg);
    });

    // Limpiar IDs auxiliares
    return data.map(op => {
      const { __excel_id, ...rest } = op;
      return rest;
    });
  }
}
