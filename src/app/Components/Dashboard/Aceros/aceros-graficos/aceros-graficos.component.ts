import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { IngresoAceros } from '../../../../models/ingreso-aceros.model';
import { SalidasAceros } from '../../../../models/salidas-aceros.model';
import { AcerosService } from '../../../../services/aceros.service';
import { ExcelAceroExportService } from '../../../../services/export/general/excel-acero-export.service';

// Interface para el resumen de diferencias
export interface DiferenciaAceros {
  proceso: string;
  tipo_acero: string;
  descripcion: string;
  cantidadIngresos: number;
  cantidadSalidas: number;
  diferencia: number;
  desglosado: boolean;
  ingresos: IngresoAceros[];
  salidas: SalidasAceros[];
}

@Component({
  selector: 'app-aceros-graficos',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './aceros-graficos.component.html',
  styleUrls: ['./aceros-graficos.component.css']
})
export class AcerosGraficosComponent implements OnInit {

  // 🔹 INGRESOS
  ingresos: IngresoAceros[] = [];
  ingresosOriginal: IngresoAceros[] = [];
  ingresosOriginalExport: IngresoAceros[] = [];

  // 🔹 SALIDAS
  salidas: SalidasAceros[] = [];
  salidasOriginal: SalidasAceros[] = [];
  salidasOriginalExport: SalidasAceros[] = [];

  // 🔹 DIFERENCIAS
  diferencias: DiferenciaAceros[] = [];

  // filtros
  fechaDesde: string = '';
  fechaHasta: string = '';
  turnoSeleccionado: string = '';
  turnos: string[] = ['DÍA', 'NOCHE'];

  constructor(private acerosService: AcerosService, private excelService: ExcelAceroExportService) {}

  ngOnInit(): void {
    const fechaISO = this.obtenerFechaLocalISO();
    this.fechaDesde = fechaISO;
    this.fechaHasta = fechaISO;
    this.turnoSeleccionado = this.obtenerTurnoActual();

    this.cargarDatos();
  }

  // ========================
  // 🔹 UTILIDADES DE FECHA
  // ========================
  obtenerFechaLocalISO(): string {
    const hoy = new Date();
    const año = hoy.getFullYear();
    const mes = (hoy.getMonth() + 1).toString().padStart(2, '0');
    const dia = hoy.getDate().toString().padStart(2, '0');
    return `${año}-${mes}-${dia}`;
  }

  obtenerFechaActualFormateada(): string {
    const hoy = new Date();
    const dia = hoy.getDate().toString().padStart(2, '0');
    const mes = (hoy.getMonth() + 1).toString().padStart(2, '0');
    const año = hoy.getFullYear();
    return `${dia}/${mes}/${año}`;
  }

  obtenerMesActual(): string {
    const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 
                  'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
    return meses[new Date().getMonth()];
  }

  obtenerTurnoActual(): string {
    const ahora = new Date();
    const hora = ahora.getHours();
    return (hora >= 7 && hora < 19) ? 'DÍA' : 'NOCHE';
  }

  // ========================
  // 🔹 CARGA DE DATOS
  // ========================
  cargarDatos(): void {
    this.acerosService.getIngresos().subscribe({
      next: (ingresosData) => {
        this.ingresosOriginal = ingresosData;
        this.ingresosOriginalExport = [...ingresosData]; // Guardar copia para exportación
        this.ingresos = [...ingresosData];
        
        this.acerosService.getSalidas().subscribe({
          next: (salidasData) => {
            this.salidasOriginal = salidasData;
            this.salidasOriginalExport = [...salidasData]; // Guardar copia para exportación
            this.salidas = [...salidasData];
            this.calcularDiferencias();
          },
          error: (err) => console.error('Error cargando salidas', err)
        });
      },
      error: (err) => console.error('Error cargando ingresos', err)
    });
  }

  // ========================
  // 🔹 CÁLCULO DE DIFERENCIAS
  // ========================
  calcularDiferencias(): void {
    const diferenciasMap = new Map<string, DiferenciaAceros>();
    
    // Procesar ingresos
    this.ingresos.forEach(ingreso => {
      const key = `${ingreso.proceso}-${ingreso.tipo_acero}-${ingreso.descripcion || ''}`;
      
      if (!diferenciasMap.has(key)) {
        diferenciasMap.set(key, {
          proceso: ingreso.proceso,
          tipo_acero: ingreso.tipo_acero,
          descripcion: ingreso.descripcion || '',
          cantidadIngresos: 0,
          cantidadSalidas: 0,
          diferencia: 0,
          desglosado: false,
          ingresos: [],
          salidas: []
        });
      }
      
      const diferencia = diferenciasMap.get(key)!;
      diferencia.cantidadIngresos += ingreso.cantidad;
      diferencia.ingresos.push(ingreso);
    });

    // Procesar salidas
    this.salidas.forEach(salida => {
      const key = `${salida.proceso}-${salida.tipo_acero}-${salida.descripcion || ''}`;
      
      if (!diferenciasMap.has(key)) {
        diferenciasMap.set(key, {
          proceso: salida.proceso,
          tipo_acero: salida.tipo_acero,
          descripcion: salida.descripcion || '',
          cantidadIngresos: 0,
          cantidadSalidas: 0,
          diferencia: 0,
          desglosado: false,
          ingresos: [],
          salidas: []
        });
      }
      
      const diferencia = diferenciasMap.get(key)!;
      diferencia.cantidadSalidas += salida.cantidad;
      diferencia.salidas.push(salida);
    });

    // Calcular diferencias finales
    diferenciasMap.forEach(diferencia => {
      diferencia.diferencia = diferencia.cantidadIngresos - diferencia.cantidadSalidas;
    });

    this.diferencias = Array.from(diferenciasMap.values());
  }

  // ========================
  // 🔹 TOGGLE DESGLOSE
  // ========================
  toggleDesglose(diferencia: DiferenciaAceros): void {
    diferencia.desglosado = !diferencia.desglosado;
  }

  // ========================
  // 🔹 FILTROS
  // ========================
  quitarFiltros(): void {
    const fechaISO = this.obtenerFechaLocalISO();
    this.fechaDesde = fechaISO;
    this.fechaHasta = fechaISO;
    this.turnoSeleccionado = this.obtenerTurnoActual();
    this.aplicarFiltros();
  }

  aplicarFiltros(): void {
    const filtros = {
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta,
      turnoSeleccionado: this.turnoSeleccionado
    };

    this.ingresos = this.filtrarDatos(this.ingresosOriginal, filtros);
    this.salidas = this.filtrarDatos(this.salidasOriginal, filtros);
    this.calcularDiferencias();
  }

  filtrarDatos<T extends { fecha: string; turno: string }>(
    datos: T[],
    filtros: any
  ): T[] {
    return datos.filter(item => {
      const fechaItem = new Date(item.fecha);
      const fechaDesde = filtros.fechaDesde ? new Date(filtros.fechaDesde) : null;
      const fechaHasta = filtros.fechaHasta ? new Date(filtros.fechaHasta) : null;

      if (fechaDesde && fechaItem < fechaDesde) return false;
      if (fechaHasta && fechaItem > fechaHasta) return false;
      if (filtros.turnoSeleccionado && item.turno !== filtros.turnoSeleccionado) return false;

      return true;
    });
  }

  // ========================
  // 🔹 MÉTODOS DE EXPORTACIÓN
  // ========================

  // Método para exportar datos COMPLETOS (sin filtrar)
  exportarAExcelExplosivos(): void {
    // Calcular diferencias completas para el stock
    const stockCompleto = this.calcularStockCompleto();
    
    this.excelService.exportarAExcelCompleto(
      this.ingresosOriginalExport,
      this.salidasOriginalExport,
      stockCompleto
    );
  }

  // Método para exportar datos FILTRADOS
  exportarAExcelExplosivosfiltro(): void {
    // Calcular diferencias filtradas para el stock
    const stockFiltrado = this.calcularStockFiltrado();
    
    this.excelService.exportarAExcelFiltrado(
      this.ingresos,
      this.salidas,
      stockFiltrado
    );
  }

  // Calcular stock completo (sin filtros)
  private calcularStockCompleto(): any[] {
    const stockMap = new Map<string, any>();
    
    // Procesar ingresos completos
    this.ingresosOriginalExport.forEach(ingreso => {
      const key = `${ingreso.proceso}-${ingreso.tipo_acero}-${ingreso.descripcion || ''}`;
      
      if (!stockMap.has(key)) {
        stockMap.set(key, {
          proceso: ingreso.proceso,
          tipo_acero: ingreso.tipo_acero,
          descripcion: ingreso.descripcion || '',
          cantidadIngresos: 0,
          cantidadSalidas: 0,
          diferencia: 0
        });
      }
      
      const stock = stockMap.get(key)!;
      stock.cantidadIngresos += ingreso.cantidad;
    });

    // Procesar salidas completas
    this.salidasOriginalExport.forEach(salida => {
      const key = `${salida.proceso}-${salida.tipo_acero}-${salida.descripcion || ''}`;
      
      if (!stockMap.has(key)) {
        stockMap.set(key, {
          proceso: salida.proceso,
          tipo_acero: salida.tipo_acero,
          descripcion: salida.descripcion || '',
          cantidadIngresos: 0,
          cantidadSalidas: 0,
          diferencia: 0
        });
      }
      
      const stock = stockMap.get(key)!;
      stock.cantidadSalidas += salida.cantidad;
    });

    // Calcular diferencias finales
    stockMap.forEach(stock => {
      stock.diferencia = stock.cantidadIngresos - stock.cantidadSalidas;
    });

    return Array.from(stockMap.values());
  }

  // Calcular stock filtrado
  private calcularStockFiltrado(): any[] {
    // Usamos las diferencias ya calculadas que están filtradas
    return this.diferencias.map(diff => ({
      proceso: diff.proceso,
      tipo_acero: diff.tipo_acero,
      descripcion: diff.descripcion,
      cantidadIngresos: diff.cantidadIngresos,
      cantidadSalidas: diff.cantidadSalidas,
      diferencia: diff.diferencia
    }));
  }
}