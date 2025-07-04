import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { FechasPlanMensualService } from '../../../services/fechas-plan-mensual.service';
import * as XLSX from 'xlsx';
import { MatDialog } from '@angular/material/dialog';
import { EmpresaService } from '../../../services/empresa.service';
import { EquipoService } from '../../../services/equipo.service';
import { TipoPerforacionService } from '../../../services/tipo-perforacion.service';
import { LoadingDialogComponent } from '../../Reutilizables/loading-dialog/loading-dialog.component';


@Component({
  selector: 'app-crear-data',
  imports: [FormsModule, CommonModule],
  templateUrl: './crear-data.component.html',
  styleUrl: './crear-data.component.css'
})
export class CrearDataComponent implements OnInit {
  modalAbierto = false;
  modalContenido: any = null;
  nuevoDato: any = {}
  formularioActivo: string = 'botones';  
  years: number[] = []; 
  meses: string[] = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  datos = [
    { nombre: 'Reporte A', year: '2024', mes: 'Enero' },
    { nombre: 'Reporte B', year: '2024', mes: 'Enero' },
    { nombre: 'Reporte C', year: '2024', mes: 'Enero' }
  ];

  constructor(
    private tipoPerforacionService: TipoPerforacionService, 
    private equipoService: EquipoService,
    private empresaService: EmpresaService,
    private FechasPlanMensualService: FechasPlanMensualService,
    public dialog: MatDialog
  ) {} // Inyecta los servicios

  ngOnInit() {
    this.generarAños();
  }

  generarAños() {
    const yearActual = new Date().getFullYear();
    for (let i = 2020; i <= yearActual; i++) {
      this.years.push(i);
    }
  }

  mostrarFormulario(formulario: string): void {
    this.formularioActivo = formulario;
  }

  buttonc = [
    {
      nombre: 'Tipo de Perforación',
      icon: 'mas.svg',
      tipo: 'Tipo de Perforación',
      datos: [],
      campos: [
        { nombre: 'nombre', label: 'Tipo de Perforación', tipo: 'text' },
        { 
          nombre: 'proceso', 
          label: 'Proceso', 
          tipo: 'select', 
          opciones: [
            'PERFORACIÓN TALADROS LARGOS',
            'PERFORACIÓN HORIZONTAL',
            'SOSTENIMIENTO',
            'SERVICIOS AUXILIARES',
            'CARGUÍO'
          ]
        }
      ]
    }, 
    {
      nombre: 'Equipo',
      icon: 'mas.svg',
      tipo: 'Equipo',
      datos: [],
      campos: [
        { nombre: 'nombre', label: 'Nombre', tipo: 'text' },
        { nombre: 'proceso', label: 'Proceso', tipo: 'text' },
        { nombre: 'codigo', label: 'Código', tipo: 'text' },
        { nombre: 'marca', label: 'Marca', tipo: 'text' },
        { nombre: 'modelo', label: 'Modelo', tipo: 'text' },
        { nombre: 'serie', label: 'Serie', tipo: 'text' },
        { nombre: 'anioFabricacion', label: 'Año de Fabricación', tipo: 'number' },
        { nombre: 'fechaIngreso', label: 'Fecha de Ingreso', tipo: 'date' },
        { nombre: 'capacidadYd3', label: 'Capacidad (Yd³)', tipo: 'number' },
        { nombre: 'capacidadM3', label: 'Capacidad (m³)', tipo: 'number' }
      ]
    },
    {
      nombre: 'Empresa',
      icon: 'mas.svg',
      tipo: 'Empresa',
      datos: [],
      campos: [
        { nombre: 'nombre', label: 'Empresa', tipo: 'text' },
      ]
    },
    {
      nombre: 'Fechas Plan Mensual',
      icon: 'mas.svg',
      tipo: 'Fechas Plan Mensual',
      datos: [],
      campos: [
        { nombre: 'mes', label: 'Mes', tipo: 'text' },
      ]
    }
    
  ];  

  cerrarModal() {
    this.modalAbierto = false;
    this.modalContenido = null;
  }

  triggerFileInput() {
    // Simula el clic en el input de archivo cuando se hace clic en el botón "Importar Excel"
    const fileInput = document.getElementById('fileInput') as HTMLInputElement;
    fileInput.click();
  }
  
  importarExcel() {
    if (this.modalContenido) {
      this.cargarExcel(this.modalContenido.nombre);
    } else {
      
    }
  }
  
  cargarExcel(nombre: string) {
    
  
    if (nombre === 'Tipo de Perforación') {
      // this.procesarExcelTipoPerforacion();
    } else if (nombre === 'Equipo') {
      this.triggerFileInput(); // Activa la selección de archivo
    } else if (nombre === 'Empresa') {
      // this.procesarExcelEmpresa();
    } else {
      
    }
  }
  
  procesarExcelEquipo(event: any) {
    const file = event.target.files[0];
  
    if (!file) {
      
      return;
    }
  
    const reader = new FileReader();
    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
  
      // Convertimos la hoja de Excel en JSON
      const excelData: any[] = XLSX.utils.sheet_to_json(sheet, { raw: false });
  
      const equipos = excelData.map(row => ({
        nombre: row["EQUIPO"] || null,
        proceso: row["PROCESO"] || null,
        codigo: row["CODIGO"] || null,
        marca: row["MARCA"] || null,
        modelo: row["MODELO"] || null,
        serie: row["SERIE"] || null,
        anioFabricacion: row["AÑO DE FABRICACIÓN"] ? Number(row["AÑO DE FABRICACIÓN"]) : null,
        fechaIngreso: this.convertirFechaExcel(row["FECHA DE INGRESO"]),
        capacidadYd3: row["CAPACIDAD (yd3)"] ? Number(row["CAPACIDAD (yd3)"]) : null,
        capacidadM3: row["CAPACIDAD (m3)"] ? Number(row["CAPACIDAD (m3)"]) : null
      }));

  
      // 🔹 Cerrar el modal antes de enviar los datos
      this.cerrarModal();
  
      // 🔹 Mostrar pantalla de carga
      const dialogRef = this.mostrarPantallaCarga();
  
      // 🔹 Enviar los datos a la API
      this.enviarEquipos(equipos)
        .then(() => {
          
          this.dialog.closeAll();
        })
        .catch((error) => {
          console.error('❌ Error al enviar datos:', error);
          this.dialog.closeAll();
        });
    };
  
    reader.readAsArrayBuffer(file);
  }
  
  convertirFechaExcel(valor: any): string | null {
    if (!valor) return null;
  
    // Si la fecha ya está en formato texto, devolverla tal cual
    if (typeof valor === "string") return valor;
  
    // Si la fecha es un número, convertirla usando XLSX
    if (typeof valor === "number") {
      const fecha = XLSX.SSF.parse_date_code(valor);
      return `${fecha.y}-${String(fecha.m).padStart(2, '0')}-${String(fecha.d).padStart(2, '0')}`;
    }
  
    return null;
  }
  
  enviarEquipos(equipos: any[]): Promise<void> {
    return new Promise((resolve, reject) => {
      const peticiones = equipos.map(nuevoRegistro => 
        this.equipoService.createEquipo(nuevoRegistro).toPromise()
      );
  
      Promise.all(peticiones)
        .then((responses) => {
          responses.forEach(data => this.modalContenido.datos.push(data));
          
          resolve();
        })
        .catch((error) => {
          console.error('❌ Error en la carga de equipos:', error);
          reject(error);
        });
    });
  }
  
  mostrarPantallaCarga() {
    this.dialog.open(LoadingDialogComponent, {
      disableClose: true
    });
  }

  
  abrirModal(button: any) {
    this.modalAbierto = true;
    this.modalContenido = button;
  
    if (button.tipo === 'Tipo de Perforación') {
      this.tipoPerforacionService.getTiposPerforacion().subscribe({
        next: (data) => {
          this.modalContenido.datos = data; // Asigna los datos recibidos
          
        },
        error: (err) => console.error('Error al cargar Tipo de Perforación:', err)
      });
    } else if (button.tipo === 'Equipo') {
      this.equipoService.getEquipos().subscribe({
        next: (data) => {
          this.modalContenido.datos = data; // Asigna los datos recibidos
          
        },
        error: (err) => console.error('Error al cargar Equipo:', err)
      });
    }else if (button.tipo === 'Empresa') {
      this.empresaService.getEmpresa().subscribe({
        next: (data) => {
          this.modalContenido.datos = data; // Asigna los datos recibidos
          
        },
        error: (err) => console.error('Error al cargar Equipo:', err)
      });
    }else if (button.tipo === 'Fechas Plan Mensual') {
      this.FechasPlanMensualService.getFechas().subscribe({
        next: (data) => {
          this.modalContenido.datos = data; // Asigna los datos recibidos
          
        },
        error: (err) => console.error('Error al cargar:', err)
      });
    }
  }

  guardarDatos() {
    if (Object.values(this.nuevoDato).some(val => val !== '')) {
      const nuevoRegistro = { ...this.nuevoDato };

      if (this.modalContenido.tipo === 'Tipo de Perforación') {
        this.tipoPerforacionService.createTipoPerforacion(nuevoRegistro).subscribe({
          next: (data) => {
            this.modalContenido.datos.push(data);
            
          },
          error: (err) => console.error('Error al guardar explosivo:', err)
        });
      } else if (this.modalContenido.tipo === 'Equipo') {
        this.equipoService.createEquipo(nuevoRegistro).subscribe({
          next: (data) => {
            this.modalContenido.datos.push(data);
            
          },
          error: (err) => console.error('Error al guardar Equipo:', err)
        });
      }else if (this.modalContenido.tipo === 'Empresa') {
        this.empresaService.createEmpresa(nuevoRegistro).subscribe({
          next: (data) => {
            this.modalContenido.datos.push(data);
            
          },
          error: (err) => console.error('Error al guardar Empresa:', err)
        });
      }else if (this.modalContenido.tipo === 'Fechas Plan Mensual') {
        this.FechasPlanMensualService.createFecha(nuevoRegistro).subscribe({
          next: (data) => {
            this.modalContenido.datos.push(data);
            
          },
          error: (err) => console.error('Error al guardar Empresa:', err)
        });
      }

      this.nuevoDato = {};
    }
  }

  eliminar(item: any): void {
    if (!item || !this.modalContenido) return;
  
    if (this.modalContenido.tipo === 'Tipo de Perforación') {
      this.tipoPerforacionService.deleteTipoPerforacion(item.id).subscribe({
        next: () => {
          this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
          
        },
        error: (err) => console.error('Error al eliminar Tipo de Perforación:', err)
      });
    } else if (this.modalContenido.tipo === 'Equipo') {
      this.equipoService.deleteEquipo(item.id).subscribe({
        next: () => {
          this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
          
        },
        error: (err) => console.error('Error al eliminar Equipo:', err)
      });
    }else if (this.modalContenido.tipo === 'Empresa') {
      this.empresaService.deleteEmpresa(item.id).subscribe({
        next: () => {
          this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
          
        },
        error: (err) => console.error('Error al eliminar accesorio:', err)
      });
    }else if (this.modalContenido.tipo === 'Fechas Plan Mensual') {
      this.FechasPlanMensualService.deleteFecha(item.id).subscribe({
        next: () => {
          this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
          
        },
        error: (err) => console.error('Error al eliminar accesorio:', err)
      });
    }
  }

  descargar(item: any): void {}
}