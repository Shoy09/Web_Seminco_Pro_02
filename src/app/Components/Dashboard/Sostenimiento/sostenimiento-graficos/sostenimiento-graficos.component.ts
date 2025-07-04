import { Component, OnInit } from '@angular/core';
import { NubeOperacion } from '../../../../models/operaciones.models';
import { OperacionService } from '../../../../services/OperacionService .service';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { NgApexchartsModule } from 'ng-apexcharts';
import { MetrosPerforadosEquipoComponent } from "../Graficos/metros-perforados-equipo/metros-perforados-equipo.component";
import { MetrosPerforadosLaborComponent } from "../Graficos/metros-perforados-labor/metros-perforados-labor.component";
import { HorometrosComponent } from "../Graficos/horometros/horometros.component";
import { MallaInstaladaEquipoComponent } from "../Graficos/malla-instalada-equipo/malla-instalada-equipo.component";
import { MallaInstaladaLaborComponent } from "../Graficos/malla-instalada-labor/malla-instalada-labor.component";
import { RendimientoDePerforacionesComponent } from "../Graficos/rendimiento-de-perforaciones/rendimiento-de-perforaciones.component";
import { DisponibilidadMecanicaEquipoComponent } from "../Graficos/disponibilidad-mecanica-equipo/disponibilidad-mecanica-equipo.component";
import { DisponibilidadMecanicaGeneralComponent } from "../Graficos/disponibilidad-mecanica-general/disponibilidad-mecanica-general.component";
import { UtilizacionGeneralComponent } from "../Graficos/utilizacion-general/utilizacion-general.component";
import { UtilizacionEquipoComponent } from "../Graficos/utilizacion-equipo/utilizacion-equipo.component";
import { GraficoEstadosComponent } from '../Graficos/grafico-estados/grafico-estados.component';
import { PromedioDeEstadosGeneralComponent } from '../Graficos/promedio-de-estados-general/promedio-de-estados-general.component';
import { Meta } from '../../../../models/meta.model';
import { MetaSostenimientoService } from '../../../../services/meta-sostenimiento.service';
import { SumaMetrosPerforadosComponent } from "../Graficos/suma-metros-perforados/suma-metros-perforados.component";
import { RendimientoPromedioComponent } from "../Graficos/rendimiento-promedio/rendimiento-promedio.component";
import { PromedioMayasComponent } from "../Graficos/promedio-mayas/promedio-mayas.component";
import { PromedioNTaladroComponent } from "../Graficos/promedio-n-taladro/promedio-n-taladro.component";
import * as XLSX from 'xlsx-js-style';

@Component({
  selector: 'app-sostenimiento-graficos',
  imports: [NgApexchartsModule, CommonModule, FormsModule, MetrosPerforadosEquipoComponent, MetrosPerforadosLaborComponent, HorometrosComponent, GraficoEstadosComponent, PromedioDeEstadosGeneralComponent, MallaInstaladaEquipoComponent, MallaInstaladaLaborComponent, RendimientoDePerforacionesComponent, DisponibilidadMecanicaEquipoComponent, DisponibilidadMecanicaGeneralComponent, UtilizacionGeneralComponent, UtilizacionEquipoComponent, SumaMetrosPerforadosComponent, RendimientoPromedioComponent, PromedioMayasComponent, PromedioNTaladroComponent],
  templateUrl: './sostenimiento-graficos.component.html',
  styleUrl: './sostenimiento-graficos.component.css'
})
export class SostenimientoGraficosComponent implements OnInit {
  datosOperaciones: NubeOperacion[] = [];
  datosOperacionesExport: NubeOperacion[] = [];
  datosGraficobarrasapiladasLargo: any[] = [];
  datosHorometros: any[] = [];
  datosGraficoEstados: any[] = [];
  datosOperacionesOriginal: NubeOperacion[] = [];
  datosGraficoMallas: any[] = [];
  datosGraficoNtaladros: any[] = [];
  RendimientoPerforacion: any[] = [];
  fechaDesde: string = '';
fechaHasta: string = '';
turnoSeleccionado: string = '';
turnos: string[] = ['DÍA', 'NOCHE'];
private todasLasMetas: Meta[] = []; 
  metasPorGrafico: { 
    [key: string]: Meta[] 
  } = {
    'METROS PERFORADOS - EQUIPO': [],
    'METROS PERFORADOS - LABOR': [],
    'ESTADOS': [],
    'ESTADOS GENERAL': [],
    'MALLA - EQUIPO': [],
    'MALLA - LABOR': [],
    'HOROMETROS': [],
    'RENDIMIENTO DE PERFORACION - EQUIPO': [],
    'DISPONIBILIDAD MECANICA - EQUIPO': [],
    'DISPONIBILIDAD MECANICA - GENERAL': [],
    'UTILIZACION - EQUIPO': [],
    'UTILIZACION - GENERAL': [],
    'SUMA DE METROS PERFORADOS': [],
    'PROMEDIO DE RENDIMIENTO': [],
    'SUMA DE TALADROS': [],
    'MALLA PROMEDIO': [],
  };


  constructor(private metaService: MetaSostenimientoService, private operacionService: OperacionService) {}
 
  ngOnInit(): void {
    const fechaISO = this.obtenerFechaLocalISO();
    this.fechaDesde = fechaISO;
    this.fechaHasta = fechaISO;
    this.turnoSeleccionado = this.obtenerTurnoActual();
  
    this.obtenerDatos();
    this.cargarMetasDesdeApi();
  }

  private cargarMetasDesdeApi(): void {
    this.metaService.getMetas().subscribe({
      next: (metas: Meta[]) => {
        if (metas && metas.length > 0) {
          this.todasLasMetas = metas;
  
          // Filtrar y agrupar las metas según el campo "grafico"
          metas.forEach(meta => {
            if (this.metasPorGrafico[meta.grafico]) {
              this.metasPorGrafico[meta.grafico].push(meta);
            }
          });
  
          // Mostrar en consola las metas por gráfico
        } else {
        }
      },
      error: (error) => {
      }
    });
  }

  obtenerTurnoActual(): string {
    const ahora = new Date();
    const hora = ahora.getHours();
  
    // Turno de día: 7:00 AM a 6:59 PM (07:00 - 18:59)
    if (hora >= 7 && hora < 19) {
      return 'DÍA';
    } else {
      // Turno de noche: 7:00 PM a 6:59 AM
      return 'NOCHE';
    }
  }  

private filtrarMetasPorMes(fechaInicio: string, fechaHasta: string): void {
  const mesSeleccionado = this.obtenerMesDeFecha(fechaInicio); // Asumiendo un mes por ahora
  const cantidadDias = this.obtenerCantidadDias(fechaInicio, fechaHasta);
  const multiplicadorTurno = this.turnoSeleccionado === '' ? 2 : 1;

  // Reiniciar metas por gráfico
  Object.keys(this.metasPorGrafico).forEach(key => {
    this.metasPorGrafico[key] = [];
  });

  this.todasLasMetas.forEach(meta => {
    if (meta.mes === mesSeleccionado && this.metasPorGrafico[meta.grafico]) {
      const metaClonada = { ...meta };
      
      // Cálculo final: objetivo * cantidad de días * multiplicador de turno
      metaClonada.objetivo = meta.objetivo * cantidadDias * multiplicadorTurno;

      this.metasPorGrafico[meta.grafico].push(metaClonada);
    }
  });
}

private obtenerCantidadDias(fechaInicio: string, fechaFin: string): number {
  const inicio = new Date(fechaInicio);
  const fin = new Date(fechaFin);
  const diffTime = Math.abs(fin.getTime() - inicio.getTime());
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 para incluir ambos días
  return diffDays;
}

  
  quitarFiltros(): void {
    const fechaISO = this.obtenerFechaLocalISO();
    this.fechaDesde = fechaISO;
    this.fechaHasta = fechaISO;
    this.turnoSeleccionado = this.obtenerTurnoActual();
  
    const filtros = {
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta,
      turnoSeleccionado: this.turnoSeleccionado
    };
  
    this.datosOperaciones = this.filtrarDatos(this.datosOperacionesOriginal, filtros);
    
    // Filtrar metas según el mes actual
    this.filtrarMetasPorMes(this.fechaDesde, this.fechaHasta);
    
    this.reprocesarTodosLosGraficos();
  }

  obtenerFechaLocalISO(): string {
    const hoy = new Date();
    const año = hoy.getFullYear();
    const mes = (hoy.getMonth() + 1).toString().padStart(2, '0'); // meses comienzan en 0
    const dia = hoy.getDate().toString().padStart(2, '0');
    return `${año}-${mes}-${dia}`;
  }
  

  aplicarFiltrosLocales(): void {
    // Crear objeto con los filtros actuales
    const filtros = {
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta,
      turnoSeleccionado: this.turnoSeleccionado
    };
  
    // Aplicar filtros a los datos ORIGINALES (this.datosOperacionesOriginal)
    const datosFiltrados = this.filtrarDatos(this.datosOperacionesOriginal, filtros);
  
    // Actualizar los datos filtrados
    this.datosOperaciones = datosFiltrados;
  
    // Filtrar metas según el mes de la fecha de inicio
    this.filtrarMetasPorMes(this.fechaDesde, this.fechaHasta);
  
    // Reprocesar los gráficos con los datos filtrados
    this.reprocesarTodosLosGraficos();
  }

private obtenerMesDeFecha(fecha: string): string {
  if (!fecha) return '';
  
  const meses = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  
  // Dividir la fecha y crear el Date objeto en UTC para evitar problemas de zona horaria
  const partes = fecha.split('-');
  const year = parseInt(partes[0], 10);
  const month = parseInt(partes[1], 10) - 1; // Restamos 1 porque los meses en Date son 0-based
  const day = parseInt(partes[2], 10);
  
  // Crear la fecha en UTC
  const date = new Date(Date.UTC(year, month, day));
  
  return meses[date.getUTCMonth()]; // Usamos getUTCMonth() para obtener el mes correcto
}
  
  filtrarDatos(datos: NubeOperacion[], filtros: any): NubeOperacion[] {
    return datos.filter(operacion => {
      const fechaOperacion = new Date(operacion.fecha);
      const fechaDesde = filtros.fechaDesde ? new Date(filtros.fechaDesde) : null;
      const fechaHasta = filtros.fechaHasta ? new Date(filtros.fechaHasta) : null;
  
      // Verificar si la fecha de operación está dentro del rango
      if (fechaDesde && fechaOperacion < fechaDesde) {
        return false;
      }
  
      if (fechaHasta && fechaOperacion > fechaHasta) {
        return false;
      }
  
      // Verificar si el turno de la operación coincide con el turno seleccionado
      if (filtros.turnoSeleccionado && operacion.turno !== filtros.turnoSeleccionado) {
        return false;
      }
  
      return true;
    });
  }
   
  
  reprocesarTodosLosGraficos(): void {
    this.prepararDatosGraficoBarrasApilada();
    this.prepararDatosHorometros();
    this.prepararDatosGraficoEstados();
    this.prepararDatosMalla();
    this.prepararDatosNtaladro();
    this.prepararDatoRendimientoPerforacion();
  }

  obtenerDatos(): void {
    this.operacionService.getOperacionesSostenimiento().subscribe({
      next: (data) => {
        this.datosOperacionesOriginal = data;
         this.datosOperacionesExport = data;
  
        // Aplicar filtros por fecha actual y turno automáticamente
        const filtros = {
          fechaDesde: this.fechaDesde,
          fechaHasta: this.fechaHasta,
          turnoSeleccionado: this.turnoSeleccionado
        };
  
        this.datosOperaciones = this.filtrarDatos(this.datosOperacionesOriginal, filtros);
  
        // Procesar datos para los gráficos
        this.prepararDatosGraficoBarrasApilada();
        this.prepararDatosHorometros();
        this.prepararDatosGraficoEstados();
        this.prepararDatosMalla();
        this.prepararDatosNtaladro();
        this.prepararDatoRendimientoPerforacion();
      },
      error: (err) => {
        console.error('❌ Error al obtener datos:', err);
      }
    });
  }

  prepararDatosGraficoBarrasApilada(): void {
    this.datosGraficobarrasapiladasLargo = this.datosOperaciones.flatMap(operacion => {
      return operacion.sostenimientos?.flatMap(perforacion => {
        return perforacion.inter_sostenimientos?.map(inter => ({
          equipo: operacion.equipo,
          codigo: operacion.codigo,
          longitud_perforacion: inter.longitud_perforacion,
          tipo_labor: perforacion.tipo_labor,
          labor: perforacion.labor,
          ntaladro: inter.ntaladro,
        })) || [];
      }) || [];
    });
  }

  prepararDatosHorometros(): void {
    this.datosHorometros = this.datosOperaciones.flatMap(operacion => 
      operacion.horometros?.map(horometro => ({
        operacionId: operacion.id,
        equipo: operacion.equipo,
        codigo: operacion.codigo,
        turno: operacion.turno,
        fecha: operacion.fecha,
        nombreHorometro: horometro.nombre,
        inicial: horometro.inicial,
        final: horometro.final,
        diferencia: horometro.final - horometro.inicial,
        EstaOP: horometro.EstaOP,
        EstaINOP: horometro.EstaINOP
      })) || []
    );
  }

  prepararDatosGraficoEstados(): void {
    this.datosGraficoEstados = this.datosOperaciones.flatMap(operacion => 
      operacion.estados?.map(estado => ({
        codigoOperacion: operacion.codigo,
        turno: operacion.turno,
        estado: estado.estado,
        codigoEstado: estado.codigo,
        hora_inicio: estado.hora_inicio,
        hora_final: estado.hora_final
      })) || []
    );
  }

  prepararDatosMalla(): void {
    this.datosGraficoMallas = this.datosOperaciones.flatMap(operacion => {
      return operacion.sostenimientos?.flatMap(perforacion => {
        return perforacion.inter_sostenimientos?.map(inter => ({
          codigo: operacion.codigo,
          malla_instalada: inter.malla_instalada,
          tipo_labor: perforacion.tipo_labor,
          labor: perforacion.labor,
        })) || [];
      }) || [];
    });
  }

  prepararDatosNtaladro(): void {
    this.datosGraficoNtaladros = this.datosOperaciones.flatMap(operacion => {
      return operacion.sostenimientos?.flatMap(perforacion => {
        return perforacion.inter_sostenimientos?.map(inter => ({
          ntaladro: inter.ntaladro,
        })) || [];
      }) || [];
    });
  }

  prepararDatoRendimientoPerforacion(): void {
    const agrupadoPorOperacion = new Map<string, {
      codigo: string;
      estados: {
        estado: string;
        codigoEstado: string;
        hora_inicio: string;
        hora_final: string;
      }[];
      perforaciones: {
        longitud_perforacion: number;
        ntaladro: number;
      }[];
    }>();
  
    for (const operacion of this.datosOperaciones) {
      const codigo = operacion.codigo;
      if (!agrupadoPorOperacion.has(codigo)) {
        agrupadoPorOperacion.set(codigo, {
          codigo,
          estados: (operacion.estados || []).map(estado => ({
            estado: estado.estado,
            codigoEstado: estado.codigo,
            hora_inicio: estado.hora_inicio,
            hora_final: estado.hora_final
          })),
          perforaciones: []
        });
      }
  
      const grupo = agrupadoPorOperacion.get(codigo)!;
  
      operacion.sostenimientos?.forEach(perforacion => {
        perforacion.inter_sostenimientos?.forEach(inter => {
          grupo.perforaciones.push({
            longitud_perforacion: inter.longitud_perforacion,
            ntaladro: inter.ntaladro
          });
        });
      });
    }
  
    // Convertimos el mapa en array final
    this.RendimientoPerforacion = Array.from(agrupadoPorOperacion.values());
  
  }
  
  exportarAExcelConHojasSeparadas(): void {
  if (!this.datosOperaciones || this.datosOperaciones.length === 0) {
    console.warn('No hay datos para exportar');
    return;
  }

  const wb: XLSX.WorkBook = XLSX.utils.book_new();

  this.datosOperaciones.forEach((operacion, index) => {
    const sheetName = `Op-${operacion.id}`.substring(0, 31);
    const datosHoja = this.prepararDatosPorOperacion(operacion);
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(datosHoja);
    
    // Aplicar estilos a la hoja
    this.aplicarEstilosAhoja(ws, datosHoja);
    
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  const fechaHoy = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Operaciones_Sostenimiento_${fechaHoy}.xlsx`);
}

aplicarEstilosAhoja(ws: XLSX.WorkSheet, datosHoja: any[][]): void {
  // Estilos base
  const estiloTituloPrincipal = {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14 },
    fill: { fgColor: { rgb: "4472C4" } }, // Azul
    alignment: { horizontal: "center" }
  };

  const estiloSubtitulo = {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
    fill: { fgColor: { rgb: "70AD47" } }, // Verde
    alignment: { horizontal: "center" }
  };

  const estiloEncabezadoTabla = {
    font: { bold: true, color: { rgb: "000000" } },
    fill: { fgColor: { rgb: "D9E1F2" } }, // Azul claro
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } }
    }
  };

  const estiloCeldaDatos = {
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } }
    }
  };

  // Aplicar estilos según el contenido
  datosHoja.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      
      // Estilo para títulos principales
      if (typeof cell === 'string' && cell.toUpperCase() === cell && cell.includes('INFORMACIÓN')) {
        ws[cellRef].s = estiloTituloPrincipal;
        if (!ws['!merges']) ws['!merges'] = [];
        ws['!merges'].push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: 8 } });
      }
      // Estilo para subtítulos (HORÓMETROS, ESTADOS, SOSTENIMIENTOS)
      else if (typeof cell === 'string' && cell.toUpperCase() === cell && 
              (cell.includes('HORÓMETROS') || cell.includes('ESTADOS') || 
               cell.includes('SOSTENIMIENTOS'))) {
        ws[cellRef].s = estiloSubtitulo;
        if (!ws['!merges']) ws['!merges'] = [];
        ws['!merges'].push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: 8 } });
      }
      // Estilo para encabezados de tabla
      else if (rowIndex > 0 && datosHoja[rowIndex-1] && datosHoja[rowIndex-1][0] && 
              typeof datosHoja[rowIndex-1][0] === 'string') {
        const cellValue = datosHoja[rowIndex-1][0].toString().toUpperCase();
        if (cellValue.includes('HORÓMETROS') || cellValue.includes('ESTADOS') || cellValue.includes('DETALLES')) {
          ws[cellRef].s = estiloEncabezadoTabla;
        }
      }
      // Estilo para datos de tabla
      else if (rowIndex > 1 && datosHoja[rowIndex-2] && datosHoja[rowIndex-2][0] && 
              typeof datosHoja[rowIndex-2][0] === 'string') {
        const cellValue = datosHoja[rowIndex-2][0].toString().toUpperCase();
        if (cellValue.includes('HORÓMETROS') || cellValue.includes('ESTADOS') || cellValue.includes('DETALLES')) {
          ws[cellRef].s = estiloCeldaDatos;
        }
      }
    });
  });

  // Ajustar el ancho de las columnas automáticamente
  const colWidths = datosHoja[0].map((_, colIndex) => {
    return {
      wch: Math.max(
        ...datosHoja.map(row => 
          row[colIndex] ? row[colIndex].toString().length + 2 : 10
        )
      )
    };
  });
  
  ws['!cols'] = colWidths;
}

prepararDatosPorOperacion(operacion: NubeOperacion): any[][] {
  const datosHoja: any[][] = [];

  // 1. Encabezado principal
  datosHoja.push(['INFORMACIÓN PRINCIPAL DE LA OPERACIÓN']);
  datosHoja.push(['ID', operacion.id]);
  datosHoja.push(['Turno', operacion.turno]);
  datosHoja.push(['Equipo', operacion.equipo]);
  datosHoja.push(['Código', operacion.codigo]);
  datosHoja.push(['Empresa', operacion.empresa]);
  datosHoja.push(['Fecha', operacion.fecha]);
  datosHoja.push(['Tipo Operación', operacion.tipo_operacion]);
  datosHoja.push(['Estado', operacion.estado]);
  datosHoja.push(['Envío', operacion.envio]);
  datosHoja.push([""]); // Espacio en blanco

  // 2. Sección de Horómetros
  if (operacion.horometros && operacion.horometros.length > 0) {
    datosHoja.push(['HORÓMETROS']);
    datosHoja.push(['Nombre', 'Inicial', 'Final', 'OP', 'INOP']);
    
    operacion.horometros.forEach(horometro => {
      datosHoja.push([
        horometro.nombre,
        horometro.inicial,
        horometro.final,
        horometro.EstaOP,
        horometro.EstaINOP
      ]);
    });
    datosHoja.push([""]); // Espacio en blanco
  }

  // 3. Sección de Estados
  if (operacion.estados && operacion.estados.length > 0) {
    datosHoja.push(['ESTADOS']);
    datosHoja.push(['Número', 'Estado', 'Código', 'Hora Inicio', 'Hora Final']);
    
    operacion.estados.forEach(estado => {
      datosHoja.push([
        estado.numero,
        estado.estado,
        estado.codigo,
        estado.hora_inicio,
        estado.hora_final
      ]);
    });
    datosHoja.push([""]); // Espacio en blanco
  }

  // 4. Sección de Sostenimientos
  if (operacion.sostenimientos && operacion.sostenimientos.length > 0) {
    datosHoja.push(['SOSTENIMIENTOS']);
    
    operacion.sostenimientos.forEach((sost, i) => {
      datosHoja.push([`SOSTENIMIENTO ${i + 1}`]);
      datosHoja.push(['Zona', sost.zona]);
      datosHoja.push(['Tipo Labor', sost.tipo_labor]);
      datosHoja.push(['Labor', sost.labor]);
      datosHoja.push(['Veta', sost.veta]);
      datosHoja.push(['Nivel', sost.nivel]);
      datosHoja.push(['Tipo Perforación', sost.tipo_perforacion]);
      
      if (sost.inter_sostenimientos && sost.inter_sostenimientos.length > 0) {
        datosHoja.push([""]);
        datosHoja.push(['DETALLES DE SOSTENIMIENTO']);
        datosHoja.push([
          'Código Actividad', 'Nivel', 'Labor', 'Sección Labor', 
          'N° Broca', 'N° Taladro', 'Longitud Perforación', 'Malla Instalada'
        ]);
        
        sost.inter_sostenimientos.forEach(inter => {
          datosHoja.push([
            inter.codigo_actividad,
            inter.nivel,
            inter.labor,
            inter.seccion_de_labor,
            inter.nbroca,
            inter.ntaladro,
            inter.longitud_perforacion,
            inter.malla_instalada
          ]);
        });
      }
      datosHoja.push([""]); // Espacio en blanco entre sostenimientos
    });
  }

  return datosHoja;
}

//EXCEL GENERALL--------------------------------------------------------------------------
prepararDatosParaExcel(datosOperaciones: NubeOperacion[]): any[] {
  const datosPlanos = [];

  for (const operacion of datosOperaciones) {
    // Datos base de la operación
    const filaBase = {
      'ID Operación': operacion.id,
      'Turno': operacion.turno,
      'Equipo': operacion.equipo,
      'Código': operacion.codigo,
      'Empresa': operacion.empresa,
      'Fecha': operacion.fecha,
      'Tipo Operación': operacion.tipo_operacion,
      'Estado': operacion.estado,
      'Envío': operacion.envio
    };

    // Procesar horómetros
    if (operacion.horometros && operacion.horometros.length > 0) {
      for (const horometro of operacion.horometros) {
        const filaHorometro = {
          ...filaBase,
          'Tipo Dato': 'HORÓMETRO',
          'Nombre Horómetro': horometro.nombre,
          'Inicial': horometro.inicial,
          'Final': horometro.final,
          'Esta OP': horometro.EstaOP,
          'Esta INOP': horometro.EstaINOP
        };
        datosPlanos.push(filaHorometro);
      }
    }

    // Procesar estados
    if (operacion.estados && operacion.estados.length > 0) {
      for (const estado of operacion.estados) {
        const filaEstado = {
          ...filaBase,
          'Tipo Dato': 'ESTADO',
          'Número Estado': estado.numero,
          'Estado': estado.estado,
          'Código Estado': estado.codigo,
          'Hora Inicio': estado.hora_inicio,
          'Hora Final': estado.hora_final
        };
        datosPlanos.push(filaEstado);
      }
    }

    // Procesar perforaciones (taladro largo)
    if (operacion.perforaciones && operacion.perforaciones.length > 0) {
      for (const perforacion of operacion.perforaciones) {
        const filaPerforacionBase = {
          ...filaBase,
          'Tipo Dato': 'PERFORACIÓN',
          'Zona': perforacion.zona,
          'Tipo Labor': perforacion.tipo_labor,
          'Labor': perforacion.labor,
          'Veta': perforacion.veta,
          'Nivel': perforacion.nivel,
          'Tipo Perforación': perforacion.tipo_perforacion
        };

        if (perforacion.inter_perforaciones && perforacion.inter_perforaciones.length > 0) {
          for (const inter of perforacion.inter_perforaciones) {
            const filaInter = {
              ...filaPerforacionBase,
              'Código Actividad': inter.codigo_actividad,
              'Nivel Inter': inter.nivel,
              'Tajo': inter.tajo,
              'N° Broca': inter.nbroca,
              'N° Taladro': inter.ntaladro,
              'N° Barras': inter.nbarras,
              'Longitud Perforación': inter.longitud_perforacion,
              'Ángulo Perforación': inter.angulo_perforacion,
              'N° Filas de Hasta': inter.nfilas_de_hasta,
              'Detalles Trabajo Realizado': inter.detalles_trabajo_realizado
            };
            datosPlanos.push(filaInter);
          }
        } else {
          datosPlanos.push(filaPerforacionBase);
        }
      }
    }

    // Procesar perforaciones horizontales
    if (operacion.perforaciones_horizontal && operacion.perforaciones_horizontal.length > 0) {
      for (const perforacion of operacion.perforaciones_horizontal) {
        const filaPerforacionHorizontalBase = {
          ...filaBase,
          'Tipo Dato': 'PERFORACIÓN HORIZONTAL',
          'Zona': perforacion.zona,
          'Tipo Labor': perforacion.tipo_labor,
          'Labor': perforacion.labor,
          'Veta': perforacion.veta,
          'Nivel': perforacion.nivel,
          'Tipo Perforación': perforacion.tipo_perforacion
        };

        if (perforacion.inter_perforaciones_horizontal && perforacion.inter_perforaciones_horizontal.length > 0) {
          for (const inter of perforacion.inter_perforaciones_horizontal) {
            const filaInterHorizontal = {
              ...filaPerforacionHorizontalBase,
              'Código Actividad': inter.codigo_actividad,
              'Nivel Inter': inter.nivel,
              'Labor Inter': inter.labor,
              'Sección Labor': inter.seccion_la_labor,
              'N° Broca': inter.nbroca,
              'N° Taladro': inter.ntaladro,
              'N° Taladros Rimados': inter.ntaladros_rimados,
              'Longitud Perforación': inter.longitud_perforacion,
              'Detalles Trabajo Realizado': inter.detalles_trabajo_realizado
            };
            datosPlanos.push(filaInterHorizontal);
          }
        } else {
          datosPlanos.push(filaPerforacionHorizontalBase);
        }
      }
    }

    // Procesar sostenimientos
    if (operacion.sostenimientos && operacion.sostenimientos.length > 0) {
      for (const sostenimiento of operacion.sostenimientos) {
        const filaSostenimientoBase = {
          ...filaBase,
          'Tipo Dato': 'SOSTENIMIENTO',
          'Zona': sostenimiento.zona,
          'Tipo Labor': sostenimiento.tipo_labor,
          'Labor': sostenimiento.labor,
          'Veta': sostenimiento.veta,
          'Nivel': sostenimiento.nivel,
          'Tipo Perforación': sostenimiento.tipo_perforacion
        };

        if (sostenimiento.inter_sostenimientos && sostenimiento.inter_sostenimientos.length > 0) {
          for (const inter of sostenimiento.inter_sostenimientos) {
            const filaInterSostenimiento = {
              ...filaSostenimientoBase,
              'Código Actividad': inter.codigo_actividad,
              'Nivel Inter': inter.nivel,
              'Labor Inter': inter.labor,
              'Sección Labor': inter.seccion_de_labor,
              'N° Broca': inter.nbroca,
              'N° Taladro': inter.ntaladro,
              'Longitud Perforación': inter.longitud_perforacion,
              'Malla Instalada': inter.malla_instalada
            };
            datosPlanos.push(filaInterSostenimiento);
          }
        } else {
          datosPlanos.push(filaSostenimientoBase);
        }
      }
    }

    // Si no hay datos relacionados, agregar solo la fila base
    if (!operacion.horometros?.length && !operacion.estados?.length && 
        !operacion.perforaciones?.length && !operacion.perforaciones_horizontal?.length &&
        !operacion.sostenimientos?.length) {
      datosPlanos.push({
        ...filaBase,
        'Tipo Dato': 'OPERACIÓN'
      });
    }
  }

  return datosPlanos;
}

exportarAExcel(): void {
  // Preparar los datos
  const datosParaExcel = this.prepararDatosParaExcel(this.datosOperacionesExport);
  
  // Crear hoja de trabajo
  const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(datosParaExcel);
  
  // Ajustar el ancho de las columnas
  const wscols = [
    {wch: 10}, // ID Operación
    {wch: 8},  // Turno
    {wch: 12}, // Equipo
    {wch: 10}, // Código
    {wch: 15}, // Empresa
    {wch: 12}, // Fecha
    {wch: 15}, // Tipo Operación
    {wch: 12}, // Estado
    {wch: 8},  // Envío
    {wch: 12}, // Tipo Dato
    {wch: 15}, // Nombre Horómetro
    {wch: 8},  // Inicial
    {wch: 8},  // Final
    {wch: 8},  // Esta OP
    {wch: 10}, // Esta INOP
    {wch: 12}, // Número Estado
    {wch: 15}, // Estado
    {wch: 15}, // Código Estado
    {wch: 12}, // Hora Inicio
    {wch: 12}, // Hora Final
    {wch: 10}, // Zona
    {wch: 15}, // Tipo Labor
    {wch: 15}, // Labor
    {wch: 10}, // Veta
    {wch: 10}, // Nivel
    {wch: 18}, // Tipo Perforación
    {wch: 15}, // Código Actividad
    {wch: 10}, // Nivel Inter
    {wch: 15}, // Labor Inter
    {wch: 15}, // Sección Labor
    {wch: 10}, // Tajo
    {wch: 8},  // N° Broca
    {wch: 10}, // N° Taladro
    {wch: 15}, // N° Taladros Rimados
    {wch: 10}, // N° Barras
    {wch: 18}, // Longitud Perforación
    {wch: 18}, // Ángulo Perforación
    {wch: 15}, // N° Filas de Hasta
    {wch: 15}, // Malla Instalada
    {wch: 25}  // Detalles Trabajo Realizado
  ];
  ws['!cols'] = wscols;

  // Definir estilos
  const headerStyle = {
    fill: {
      patternType: 'solid',
      fgColor: { rgb: '4472C4' } // Azul corporativo
    },
    font: {
      name: 'Arial',
      sz: 11,
      bold: true,
      color: { rgb: 'FFFFFF' } // Texto blanco
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true
    },
    border: {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    }
  };

  const dataStyle = {
    font: {
      name: 'Arial',
      sz: 10
    },
    border: {
      top: { style: 'thin', color: { rgb: 'D9D9D9' } },
      bottom: { style: 'thin', color: { rgb: 'D9D9D9' } },
      left: { style: 'thin', color: { rgb: 'D9D9D9' } },
      right: { style: 'thin', color: { rgb: 'D9D9D9' } }
    }
  };

  // Aplicar estilos a los encabezados
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z1');
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
    if (!ws[cellAddress]) continue;
    ws[cellAddress].s = headerStyle;
  }

  // Aplicar estilos a los datos
  for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellAddress]) continue;
      
      // Estilo base para datos
      ws[cellAddress].s = dataStyle;
      
      // Formato condicional por tipo de dato
      const tipoDato = ws['J' + (R + 1)]?.v; // Columna J es 'Tipo Dato'
      
      if (tipoDato === 'HORÓMETRO') {
        ws[cellAddress].s.fill = { patternType: 'solid', fgColor: { rgb: 'E2EFDA' } }; // Verde claro
      } else if (tipoDato === 'ESTADO') {
        ws[cellAddress].s.fill = { patternType: 'solid', fgColor: { rgb: 'FFF2CC' } }; // Amarillo claro
      } else if (tipoDato === 'PERFORACIÓN') {
        ws[cellAddress].s.fill = { patternType: 'solid', fgColor: { rgb: 'DDEBF7' } }; // Azul claro
      } else if (tipoDato === 'PERFORACIÓN HORIZONTAL') {
        ws[cellAddress].s.fill = { patternType: 'solid', fgColor: { rgb: 'F8CBAD' } }; // Naranja claro
      } else if (tipoDato === 'SOSTENIMIENTO') {
        ws[cellAddress].s.fill = { patternType: 'solid', fgColor: { rgb: 'EAD1DC' } }; // Lila claro
      }
      
      // Formato para números
      if (typeof ws[cellAddress].v === 'number') {
        ws[cellAddress].s.numFmt = '0.00';
        ws[cellAddress].s.alignment = { horizontal: 'right' };
      }
      
      // Formato para fechas
      if (cellAddress.includes('Fecha') || cellAddress.includes('Hora')) {
        ws[cellAddress].s.numFmt = 'dd/mm/yyyy hh:mm';
      }
    }
  }

  // Congelar la primera fila (encabezados)
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomRight' };

  // Añadir filtros a los encabezados
  ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };

  // Crear libro de trabajo
  const wb: XLSX.WorkBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'datosOperacionesExport');
  
  // Exportar el archivo
  const fechaHoy = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Operaciones_Sostenimiento_${fechaHoy}.xlsx`, { compression: true });
}

}
