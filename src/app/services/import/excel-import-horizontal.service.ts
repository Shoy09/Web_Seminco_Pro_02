import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeCheckListOperacion, NubeHorometros, NubeInterPerforacionHorizontal } from '../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelImportHorizontalService {

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
            const horometro: Omit<NubeHorometros, 'id' | 'operacion_id'> = {
              nombre,
              inicial: Number(row[`Horómetro ${nombre} - Inicial`] || 0),
              final: Number(row[`Horómetro ${nombre} - Final`] || 0),
              EstaOP: row[`Horómetro ${nombre} - Operativo`] === 'Sí' ? 1 : 0,
              EstaINOP: row[`Horómetro ${nombre} - Operativo`] === 'No' ? 1 : 0
            };
            op.horometros.push(horometro);
          }
        });
      }
    });

    // Mapear checklist
    checklistData.forEach(row => {
      const op = data.find(op => op.__excel_id === row['ID Operación']);
      if (op && !row['Mensaje']) {
        const check: Omit<NubeCheckListOperacion, 'id' | 'operacion_id'> = {
          descripcion: row['Descripción'] || '',
          decision: row['Decisión'] === 'Aprobado' ? 1 : 0,
          observacion: row['Observación'] || '',
          categoria: row['Categoría'] || ''
        };
        op.checklists.push(check);
      }
    });

    // Mapear estados y perforaciones/interperforaciones horizontales
    estadosData.forEach(row => {
      const op = data.find(op => op.__excel_id === row['ID Operación']);
      if (!op || row['Mensaje']) return;

      let estado = op.estados.find((e: any) => e.numero === Number(row['Número Estado'] || 0));
      if (!estado) {
        estado = {
          numero: Number(row['Número Estado'] || 0),
          estado: row['Estado'] || '',
          codigo: row['Código Estado'] || '',
          hora_inicio: row['Hora Inicio'] || '',
          hora_final: row['Hora Final'] || '',
          perforaciones_horizontal: []
        };
        op.estados.push(estado);
      }

      const zona = (row['Perf. Horizontal - Zona'] || '').toString().trim();
      const tipoLabor = (row['Perf. Horizontal - Tipo Labor'] || '').toString().trim();

      let perf = estado.perforaciones_horizontal.find((p: any) =>
        p.zona?.toString().trim() === zona && p.tipo_labor?.toString().trim() === tipoLabor
      );

      if (!perf && zona) {
        perf = {
          zona,
          tipo_labor: tipoLabor,
          labor: row['Perf. Horizontal - Labor'] || '',
          veta: row['Perf. Horizontal - Veta'] || '',
          nivel: row['Perf. Horizontal - Nivel'] || '',
          tipo_perforacion: row['Perf. Horizontal - Tipo Perforación'] || '',
          observacion: row['Perf. Horizontal - Observación'] || '',
          inter_perforaciones_horizontal: []
        };
        estado.perforaciones_horizontal.push(perf);
      }

      if (perf) {
        const inter: Omit<NubeInterPerforacionHorizontal, 'id' | 'perforacionhorizontal_id'> = {
          codigo_actividad: row['Ejecutado - Código Actividad'] || '',
          nivel: row['Ejecutado - Nivel'] || '',
          labor: row['Ejecutado - Labor'] || '',
          seccion_la_labor: row['Ejecutado - Sección'] || '',
          nbroca: Number(row['Ejecutado - N° Broca'] || 0),
          ntaladro: Number(row['Ejecutado - N° Taladro'] || 0),
          material: row['Ejecutado - Material'] || '',
          ntaladros_rimados: Number(row['Ejecutado - N° Taladros Rimados'] || 0),
          longitud_perforacion: Number(row['Ejecutado - Longitud'] || 0),
          detalles_trabajo_realizado: row['Ejecutado - Detalles'] || '',
          metros_perforados: Number(row['Ejecutado - Metros perforados'] || 0)
        };
        perf.inter_perforaciones_horizontal.push(inter);
      }
    });

    // 🔥 limpiar ids y normalizar antes de enviar
    return data.map(op => {
      const { __excel_id, ...rest } = op;
      return {
        operacion: {
          ...rest.operacion,
          envio: 0,
        },
        horometros: rest.horometros.map((h: any) => ({
          nombre: h.nombre,
          inicial: Number(h.inicial) || 0,
          final: Number(h.final) || 0,
          EstaOP: Number(h.EstaOP) || 0,
          EstaINOP: Number(h.EstaINOP) || 0,
        })),
        checklists: rest.checklists.map((c: any) => ({
          descripcion: c.descripcion || '',
          decision: Number(c.decision) || 0,
          observacion: c.observacion || '',
          categoria: c.categoria || ''
        })),
        // Dentro de la parte de mapeo de estados
estados: rest.estados.map((e: any) => ({
  numero: Number(e.numero) || 0,
  estado: e.estado || '',
  codigo: e.codigo || '',
  hora_inicio: e.hora_inicio || '',
  hora_final: e.hora_final || '',
  perforaciones_horizontales: e.perforaciones_horizontal.map((p: any) => ({
    zona: p.zona || '',
    tipo_labor: p.tipo_labor || '',
    labor: p.labor || '',
    veta: p.veta || '',
    nivel: p.nivel || '',
    tipo_perforacion: p.tipo_perforacion || '',
    observacion: p.observacion || '',
    inter_perforaciones: p.inter_perforaciones_horizontal.map((ip: any) => ({
      codigo_actividad: ip.codigo_actividad || '',
      nivel: ip.nivel || '',
      labor: ip.labor || '',
      seccion_la_labor: ip.seccion_la_labor || '',
      nbroca: Number(ip.nbroca) || 0,
      ntaladro: Number(ip.ntaladro) || 0,
      material: ip.material || '',
      ntaladros_rimados: Number(ip.ntaladros_rimados) || 0,
      longitud_perforacion: Number(ip.longitud_perforacion) || 0,
      detalles_trabajo_realizado: ip.detalles_trabajo_realizado || '',
      metros_perforados: Number(ip.metros_perforados) || 0
    }))
  }))
}))

      };
    });
  }
}
