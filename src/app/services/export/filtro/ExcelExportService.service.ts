import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion } from '../../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelExportServiceLargoFiltro {
exportOperacionesToExcel(datosOperaciones: NubeOperacion[], fileName: string) {
  // Filtrar solo operaciones con estado "Cerrado"
  const operacionesCerradas = datosOperaciones.filter(op => 
    op.estado?.toLowerCase() === 'cerrado' // Case-insensitive
  );

  // Preparar datos para cada hoja (usando solo operaciones cerradas)
  const ejecutadoData = this.prepareEjecutadoData(operacionesCerradas);
  const estadosData = this.prepareEstadosData(operacionesCerradas);
  const checklistData = this.prepareChecklistData(operacionesCerradas);

  // Resto del código permanece igual...
  const wb: XLSX.WorkBook = XLSX.utils.book_new();
  const ejecutadoWS = XLSX.utils.json_to_sheet(ejecutadoData);
  const estadosWS = XLSX.utils.json_to_sheet(estadosData);
  const checklistWS = XLSX.utils.json_to_sheet(checklistData);

  this.adjustColumnWidth(ejecutadoWS, ejecutadoData);
  this.adjustColumnWidth(estadosWS, estadosData);
  this.adjustColumnWidth(checklistWS, checklistData);

  XLSX.utils.book_append_sheet(wb, ejecutadoWS, 'EJECUTADOTL');
  XLSX.utils.book_append_sheet(wb, estadosWS, 'ESTADOSTL');
  XLSX.utils.book_append_sheet(wb, checklistWS, 'CHECK LISTTL');

  XLSX.writeFile(wb, `${fileName}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

private prepareEjecutadoData(operaciones: NubeOperacion[]): any[] {
  const data: any[] = [];
  
  operaciones.forEach(op => {
    // Crear objeto base para la operación
    const rowData: any = {
      'ID Operación': op.id,
      'Turno': op.turno,
      'Equipo': op.equipo,
      'Código': op.codigo,
      'Empresa': op.empresa,
      'Fecha': op.fecha,
      'Tipo Operación': op.tipo_operacion,
      'Estado': op.estado,
    };

    // Procesar horómetros si existen
    if (op.horometros && op.horometros.length > 0) {
      op.horometros.forEach(horometro => {
        const nombreNormalizado = horometro.nombre.replace(/\s+/g, '_'); // Normalizar nombre para nombres compuestos
        
        // Agregar columnas para cada horómetro (sin el ID)
        rowData[`Horómetro ${nombreNormalizado} - Inicial`] = horometro.inicial;
        rowData[`Horómetro ${nombreNormalizado} - Final`] = horometro.final;
        rowData[`Diferencia ${nombreNormalizado}`] = horometro.final - horometro.inicial;
        
        // Determinar el estado operativo según las reglas
        const opValue = horometro.EstaOP;
        const inopValue = horometro.EstaINOP;
        let estadoOperativo;
        
        if (opValue && !inopValue) {
          estadoOperativo = 'Sí';
        } else if (!opValue && inopValue) {
          estadoOperativo = 'No';
        } else {
          estadoOperativo = 'Sin definir';
        }
        
        rowData[`Horómetro ${nombreNormalizado} - Operativo`] = estadoOperativo;
      });
    }
const fechaMina = this.calcularFechaMina(op.fecha, op.turno);
    rowData['Fecha_Mina'] = fechaMina;
    data.push(rowData);
  });
  
  return data;
}

  private prepareEstadosData(operaciones: NubeOperacion[]): any[] {
    const data: any[] = [];
    
    operaciones.forEach(op => {
       const fechaMina = this.calcularFechaMina(op.fecha, op.turno);
      if (op.estados && op.estados.length > 0) {
        op.estados.forEach(estado => {
          // Datos base del estado
          const estadoBase = {
            'ID Operación': op.id,
            'ID Estado': estado.id,
            'Número Estado': estado.numero,
            'Estado': estado.estado,
            'Código Estado': estado.codigo,
            'Hora Inicio': estado.hora_inicio,
            'Hora Final': estado.hora_final,
            // Campos de perforación (inicialmente vacíos)
            'Perf. - Zona': '',
            'Perf. - Tipo Labor': '',
            'Perf. - Labor': '',
            'Perf. - Veta': '',
            'Perf. - Nivel': '',
            'Perf. - Tipo Perforación': '',
            'Perf. - Observación': '',
            // Campos de interperforación (inicialmente vacíos)
            'Ejecutado - Código Actividad': '',
            'Ejecutado - Nivel': '',
            'Ejecutado - Tajo': '',
            'Ejecutado - N° Broca': '',
            'Ejecutado - N° Taladro': '',
            'Ejecutado - Material': '',
            'Ejecutado - N° Barras': '',
            'Ejecutado - Longitud': '',
            'Ejecutado - Ángulo': '',
            'Ejecutado - N° Filas': '',
            'Ejecutado - Detalles': '',
            'Metros perforados': 0,
            'Fecha_Mina': fechaMina, 
          };

          // Procesar perforaciones taladro largo si existen
          if (estado.perforaciones_taladro_largo?.length) {
          estado.perforaciones_taladro_largo.forEach(perf => {
            const rowWithPerf = {
              ...estadoBase,
                'Perf. - Zona': perf.zona,
                'Perf. - Tipo Labor': perf.tipo_labor,
                'Perf. - Labor': perf.labor,
                'Perf. - Veta': perf.veta,
                'Perf. - Nivel': perf.nivel,
                'Perf. - Tipo Perforación': perf.tipo_perforacion,
                'Perf. - Observación': perf.observacion || ''
              };

              // Procesar interperforaciones si existen
              if (perf.inter_perforaciones?.length) {
              perf.inter_perforaciones.forEach(inter => {
                const nTaladro = inter.ntaladro || 0;
                const longitud = inter.longitud_perforacion || 0;
                const metrosPerforados = nTaladro * longitud;

                data.push({
                  ...rowWithPerf,
                    'Ejecutado - Código Actividad': inter.codigo_actividad,
                    'Ejecutado - Nivel': inter.nivel,
                    'Ejecutado - Tajo': inter.tajo,
                    'Ejecutado - N° Broca': inter.nbroca,
                    'Ejecutado - N° Taladro': inter.ntaladro,
                    'Ejecutado - Material': inter.material,
                    'Ejecutado - N° Barras': inter.nbarras,
                    'Ejecutado - Longitud': inter.longitud_perforacion,
                    'Ejecutado - Ángulo': inter.angulo_perforacion,
                    'Ejecutado - N° Filas': inter.nfilas_de_hasta,
                    'Ejecutado - Detalles': inter.detalles_trabajo_realizado,
                    'Metros perforados': metrosPerforados
                  });
                });
              } else {
                data.push(rowWithPerf);
              }
            });
          } else {
            data.push(estadoBase);
          }
        });
      } else {
        // Si no hay estados, exportar solo datos básicos de operación
        data.push({
          'ID Operación': op.id,
          'Mensaje': 'No hay estados registrados para esta operación'
        });
      }
    });
    
    return data;
  }

  private calcularFechaMina(fechaOriginal: string, turno: string): string {
  if (!fechaOriginal) return '';
  
  // Si el turno es "Noche", sumar un día a la fecha original
  if (turno?.toLowerCase() === 'noche') {
    const fecha = new Date(fechaOriginal);
    fecha.setDate(fecha.getDate() + 1);
    return fecha.toISOString().split('T')[0];
  }
  
  // Para cualquier otro caso (incluyendo turno "Dia"), usar la fecha original
  return fechaOriginal.split('T')[0];
}
  private prepareChecklistData(operaciones: NubeOperacion[]): any[] {
    const data: any[] = [];
    
    operaciones.forEach(op => {
      if (op.checklists && op.checklists.length > 0) {
        op.checklists.forEach(check => {
          data.push({
            'ID Operación': op.id,
            'Descripción': check.descripcion,
            'Decisión': check.decision === 1 ? 'Aprobado' : 'Rechazado',
            'Observación': check.observacion,
            'Categoría': check.categoria
          });
        });
      } else {
        // Si no hay checklist, exportar solo datos básicos de operación
        data.push({
          'ID Operación': op.id,
          'Mensaje': 'No hay checklist registrado para esta operación'
        });
      }
    });
    
    return data;
  }

  private adjustColumnWidth(worksheet: XLSX.WorkSheet, data: any[]) {
    if (!data || data.length === 0) return;

    const columnWidths: XLSX.ColInfo[] = [];
    const headers = Object.keys(data[0]);

    // Calcular el ancho máximo para cada columna
    headers.forEach((header, i) => {
      // Ancho mínimo basado en el encabezado
      let maxWidth = header.length * 1.2;
      
      // Verificar el contenido de cada fila
      data.forEach(row => {
        const value = row[header];
        if (value !== undefined && value !== null) {
          const length = value.toString().length;
          if (length > maxWidth) {
            maxWidth = length * 1.1;
          }
        }
      });

      // Limitar entre 10 y 50 caracteres de ancho
      const width = Math.min(Math.max(maxWidth, 10), 50);
      columnWidths.push({ wch: width });
    });

    worksheet['!cols'] = columnWidths;
  }
}