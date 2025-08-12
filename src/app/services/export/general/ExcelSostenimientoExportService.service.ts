import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion } from '../../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelSostenimientoExportService {
  exportOperacionesToExcel(datosOperacionesExport: NubeOperacion[], fileName: string) {

      const operacionesCerradas = datosOperacionesExport.filter(op => 
    op.estado?.toLowerCase() === 'cerrado' // Case-insensitive
  );

    // Preparar datos para cada hoja
    const ejecutadoData = this.prepareEjecutadoData(operacionesCerradas);
    const estadosData = this.prepareEstadosData(operacionesCerradas);
    const checklistData = this.prepareChecklistData(operacionesCerradas);

    // Crear un nuevo libro de trabajo
    const wb: XLSX.WorkBook = XLSX.utils.book_new();

    // Añadir hojas al libro
    const ejecutadoWS = XLSX.utils.json_to_sheet(ejecutadoData);
    const estadosWS = XLSX.utils.json_to_sheet(estadosData);
    const checklistWS = XLSX.utils.json_to_sheet(checklistData);

    // Ajustar el ancho de las columnas
    this.adjustColumnWidth(ejecutadoWS, ejecutadoData);
    this.adjustColumnWidth(estadosWS, estadosData);
    this.adjustColumnWidth(checklistWS, checklistData);

    XLSX.utils.book_append_sheet(wb, ejecutadoWS, 'Ejecutado');
    XLSX.utils.book_append_sheet(wb, estadosWS, 'Estados');
    XLSX.utils.book_append_sheet(wb, checklistWS, 'Checklist');

    // Exportar el archivo
    XLSX.writeFile(wb, `${fileName}_Sostenimiento_${new Date().toISOString().slice(0,10)}.xlsx`);
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

    data.push(rowData);
  });
  
  return data;
}

private prepareEstadosData(operaciones: NubeOperacion[]): any[] {
  const data: any[] = [];
  
  operaciones.forEach(op => {
    if (op.estados?.length) {
      op.estados.forEach(estado => {
        // Datos base del estado (con nueva columna al final)
        const estadoBase = {
          'ID Operación': op.id,
          'ID Estado': estado.id,
          'Número Estado': estado.numero,
          'Estado': estado.estado,
          'Código Estado': estado.codigo,
          'Hora Inicio': estado.hora_inicio,
          'Hora Final': estado.hora_final,
          'Sost. - Zona': '',
          'Sost. - Tipo Labor': '',
          'Sost. - Labor': '',
          'Sost. - Veta': '',
          'Sost. - Nivel': '',
          'Sost. - Tipo Perforación': '',
          'Sost. - Observación': '',
          'Ejecutado - Código Actividad': '',
          'Ejecutado - Nivel': '',
          'Ejecutado - Labor': '',
          'Ejecutado - Sección': '',
          'Ejecutado - N° Broca': '',
          'Ejecutado - N° Taladro': '',
          'Ejecutado - Material': '',
          'Ejecutado - Longitud': '',
          'Ejecutado - Malla Instalada': '',
          'Ejecutado - Detalles': '',
          'Metros perforados': 0 // Nueva columna
        };

        // Procesar sostenimientos si existen
        if (estado.sostenimientos?.length) {
          estado.sostenimientos.forEach(sost => {
            const rowWithSost = {
              ...estadoBase,
              'Sost. - Zona': sost.zona || '',
              'Sost. - Tipo Labor': sost.tipo_labor || '',
              'Sost. - Labor': sost.labor || '',
              'Sost. - Veta': sost.veta || '',
              'Sost. - Nivel': sost.nivel || '',
              'Sost. - Tipo Perforación': sost.tipo_perforacion || '',
              'Sost. - Observación': sost.observacion || ''
            };

            // Procesar inter-sostenimientos si existen
            if (sost.inter_sostenimientos?.length) {
              sost.inter_sostenimientos.forEach(inter => {
                const nTaladro = inter.ntaladro || 0;
                const longitud = inter.longitud_perforacion || 0;
                const metrosPerforados = nTaladro * longitud;

                data.push({
                  ...rowWithSost,
                  'Ejecutado - Código Actividad': inter.codigo_actividad || '',
                  'Ejecutado - Nivel': inter.nivel || '',
                  'Ejecutado - Labor': inter.labor || '',
                  'Ejecutado - Sección': inter.seccion_de_labor || '',
                  'Ejecutado - N° Broca': inter.nbroca || '',
                  'Ejecutado - N° Taladro': nTaladro,
                  'Ejecutado - Material': inter.material,
                  'Ejecutado - Longitud': longitud,
                  'Ejecutado - Malla Instalada': inter.malla_instalada || '',
                  'Ejecutado - Detalles': inter.detalles_trabajo_realizado || '',
                  'Metros perforados': metrosPerforados
                });
              });
            } else {
              data.push(rowWithSost);
            }
          });
        } else {
          data.push(estadoBase);
        }
      });
    } else {
      data.push({
        'ID Operación': op.id,
        'Mensaje': 'No hay estados registrados para esta operación'
      });
    }
  });
  
  return data;
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

    headers.forEach((header, i) => {
      let maxWidth = header.length * 1.2;
      
      data.forEach(row => {
        const value = row[header];
        if (value !== undefined && value !== null) {
          const length = value.toString().length;
          if (length > maxWidth) {
            maxWidth = length * 1.1;
          }
        }
      });

      const width = Math.min(Math.max(maxWidth, 10), 50);
      columnWidths.push({ wch: width });
    });

    worksheet['!cols'] = columnWidths;
  }
}