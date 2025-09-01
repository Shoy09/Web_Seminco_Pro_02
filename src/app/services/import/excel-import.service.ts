import { Injectable } from '@angular/core';
import { NubeCheckListOperacion, NubeHorometros, NubeInterPerforacionTaladroLargo } from '../../models/operaciones.models';
import * as XLSX from 'xlsx';

@Injectable({
  providedIn: 'root'
})
export class ExcelImportService {

  constructor() {}

  async importOperacionesFromExcel(file: File): Promise<any[]> {
    const data: any[] = [];

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

    const ejecutadoSheet = workbook.Sheets['EJECUTADOTL'];
    const estadosSheet = workbook.Sheets['ESTADOSTL'];
    const checklistSheet = workbook.Sheets['CHECK LISTTL'];

    const ejecutadoData = XLSX.utils.sheet_to_json<any>(ejecutadoSheet);
    const estadosData = XLSX.utils.sheet_to_json<any>(estadosSheet);
    const checklistData = XLSX.utils.sheet_to_json<any>(checklistSheet);

    // Mapear operaciones básicas
    ejecutadoData.forEach(row => {
      const idOp = row['ID Operación'];
      if (!data.find(op => op.__excel_id === idOp)) {
        data.push({
          __excel_id: idOp, // ⚠️ solo para agrupar, luego lo quitamos
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

    // Mapear estados y perforaciones/interperforaciones
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
          perforaciones: []
        };
        op.estados.push(estado);
      }

      const zona = (row['Perf. - Zona'] || '').toString().trim();
      const tipoLabor = (row['Perf. - Tipo Labor'] || '').toString().trim();

      let perf = estado.perforaciones.find((p: any) =>
        p.zona?.toString().trim() === zona && p.tipo_labor?.toString().trim() === tipoLabor
      );

      if (!perf && zona) {
        perf = {
          zona,
          tipo_labor: tipoLabor,
          labor: row['Perf. - Labor'] || '',
          veta: row['Perf. - Veta'] || '',
          nivel: row['Perf. - Nivel'] || '',
          tipo_perforacion: row['Perf. - Tipo Perforación'] || '',
          observacion: row['Perf. - Observación'] || '',
          inter_perforaciones: []
        };
        estado.perforaciones.push(perf);
      }

      if (perf) {
        const inter: Omit<NubeInterPerforacionTaladroLargo, 'id' | 'perforaciontaladrolargo_id'> = {
          codigo_actividad: row['Ejecutado - Código Actividad'] || '',
          nivel: row['Ejecutado - Nivel'] || '',
          tajo: row['Ejecutado - Tajo'] || '',
          nbroca: Number(row['Ejecutado - N° Broca'] || 0),
          ntaladro: Number(row['Ejecutado - N° Taladro'] || 0),
          material: row['Ejecutado - Material'] || '',
          nbarras: Number(row['Ejecutado - N° Barras'] || 0),
          longitud_perforacion: Number(row['Ejecutado - Longitud'] || 0),
          angulo_perforacion: Number(row['Ejecutado - Ángulo'] || 0),
          nfilas_de_hasta: row['Ejecutado - N° Filas'] || '',
          detalles_trabajo_realizado: row['Ejecutado - Detalles'] || ''
        };
        perf.inter_perforaciones.push(inter);
      }
    });

    // 🔥 limpiar ids usados solo para agrupar
    // 🔥 limpiar ids y normalizar antes de enviar
return data.map(op => {
  const { __excel_id, ...rest } = op;

  return {
    operacion: {
      ...rest.operacion,
      envio: 0, // aseguramos default
    },
    horometros: rest.horometros.map((h: { nombre: any; inicial: any; final: any; EstaOP: any; EstaINOP: any; }) => ({
      nombre: h.nombre,
      inicial: Number(h.inicial) || 0,
      final: Number(h.final) || 0,
      EstaOP: Number(h.EstaOP) || 0,
      EstaINOP: Number(h.EstaINOP) || 0,
    })),
    checklists: rest.checklists.map((c: { descripcion: any; decision: any; observacion: any; categoria: any; }) => ({
      descripcion: c.descripcion || '',
      decision: Number(c.decision) || 0,
      observacion: c.observacion || '',
      categoria: c.categoria || ''
    })),
    estados: rest.estados.map((e: { numero: any; estado: any; codigo: any; hora_inicio: any; hora_final: any; perforaciones: any[]; }) => ({
      numero: Number(e.numero) || 0,
      estado: e.estado || '',
      codigo: e.codigo || '',
      hora_inicio: e.hora_inicio || '',
      hora_final: e.hora_final || '',
      perforaciones: e.perforaciones.map((p: { zona: any; tipo_labor: any; labor: any; veta: any; nivel: any; tipo_perforacion: any; observacion: any; inter_perforaciones: any[]; }) => ({
        zona: p.zona || '',
        tipo_labor: p.tipo_labor || '',
        labor: p.labor || '',
        veta: p.veta || '',
        nivel: p.nivel || '',
        tipo_perforacion: p.tipo_perforacion || '',
        observacion: p.observacion || '',
        inter_perforaciones: p.inter_perforaciones.map((ip: { codigo_actividad: any; nivel: any; tajo: any; nbroca: any; ntaladro: any; material: any; nbarras: any; longitud_perforacion: any; angulo_perforacion: any; nfilas_de_hasta: any; detalles_trabajo_realizado: any; }) => ({
          codigo_actividad: ip.codigo_actividad || '',
          nivel: ip.nivel || '',
          tajo: ip.tajo || '',
          nbroca: Number(ip.nbroca) || 0,
          ntaladro: Number(ip.ntaladro) || 0,
          material: ip.material || '',
          nbarras: Number(ip.nbarras) || 0,
          longitud_perforacion: Number(ip.longitud_perforacion) || 0,
          angulo_perforacion: Number(ip.angulo_perforacion) || 0,
          nfilas_de_hasta: ip.nfilas_de_hasta || '',
          detalles_trabajo_realizado: ip.detalles_trabajo_realizado || ''
        }))
      }))
    }))
  };
});

  }
}
