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
import { ToneladasService } from '../../../services/toneladas.service';
import { OrigenDestinoService } from '../../../services/origen-destino.service';
import { JefeGuardiaAceroService } from '../../../services/jefe-guardia-acero.service';
import { OperadorAceroService } from '../../../services/operador-acero.service';
import { ProcesoAceroService } from '../../../services/proceso-acero.service';


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
    editando: boolean = false;
indiceEditando: number = -1;
datoOriginal: any = null;
  meses: string[] = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  datos = [
    { nombre: 'Reporte A', year: '2024', mes: 'Enero' },
    { nombre: 'Reporte B', year: '2024', mes: 'Enero' },
    { nombre: 'Reporte C', year: '2024', mes: 'Enero' }
  ];

  constructor(
    private tipoPerforacionService: TipoPerforacionService, 
    private equipoService: EquipoService,
    private origenDestinoService: OrigenDestinoService,
    private empresaService: EmpresaService,
    private FechasPlanMensualService: FechasPlanMensualService,
    private toneladasService: ToneladasService,
    public dialog: MatDialog,
        private procesoAceroService: ProcesoAceroService,
private jefeGuardiaAceroService: JefeGuardiaAceroService,
    private operadorAceroService: OperadorAceroService,
    
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
    },
    { 
      nombre: 'permitido_medicion', 
      label: '¿Permitido para medición?', 
      tipo: 'select', 
      opciones: [
        'SI',
        'NO'
      ]
    },
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
    },
    {
  nombre: 'Toneladas',
  icon: 'mas.svg',
  tipo: 'Toneladas',
  datos: [],
  campos: [
    { nombre: 'fecha', label: 'Fecha', tipo: 'date' },
    { 
          nombre: 'turno', 
          label: 'Turno', 
          tipo: 'select', 
          opciones: [
            'Dia',
            'Noche'
          ]
        },
    { nombre: 'zona', label: 'Zona', tipo: 'text' },
    { nombre: 'tipo', label: 'Tipo', tipo: 'text' },
    { nombre: 'labor', label: 'Labor', tipo: 'text' },
    { nombre: 'toneladas', label: 'Toneladas', tipo: 'number' }
  ]
},
{
  nombre: 'Origen/Destino',
  icon: 'mas.svg',
  tipo: 'OrigenDestino',
  datos: [],
  campos: [
    { 
      nombre: 'operacion', 
      label: 'Operación', 
      tipo: 'text', 
      valorPorDefecto: 'CARGUÍO'
    },
    { 
      nombre: 'tipo', 
      label: 'Tipo', 
      tipo: 'select',
      opciones: ['Origen', 'Destino']
    },
    { nombre: 'nombre', label: 'Nombre', tipo: 'text' }
  ]
},
{
  nombre: 'Acero',
  icon: 'mas.svg',
  tipo: 'Acero',
  datos: [],
  campos: [
    { 
      nombre: 'proceso', 
      label: 'Proceso', 
      tipo: 'select', 
      opciones: [
        'PERFORACIÓN TALADROS LARGOS',
        'PERFORACIÓN HORIZONTAL',
        'SOSTENIMIENTO',
      ]
    },
    { nombre: 'tipo_acero', label: 'Tipo de Acero', tipo: 'text' },
    { nombre: 'descripcion', label: 'Descripción', tipo: 'text' },
    { nombre: 'precio', label: 'Precio', tipo: 'number' }
  ]
},
{
  nombre: 'Jefe Guardia Acero',
  icon: 'mas.svg',
  tipo: 'JefeGuardiaAcero',
  datos: [],
  campos: [
    { nombre: 'jefe_de_guardia', label: 'Nombre', tipo: 'text' },
    { 
      nombre: 'turno', 
      label: 'Turno', 
      tipo: 'select',
      opciones: ['DIA', 'NOCHE']
    },
    { 
      nombre: 'activo', 
      label: 'Activo', 
      tipo: 'select',
      opciones: ['SI', 'NO']
    }
  ]
},
{
  nombre: 'Operador Acero',
  icon: 'mas.svg',
  tipo: 'OperadorAcero',
  datos: [],
  campos: [
    { nombre: 'operador', label: 'Nombre', tipo: 'text' },
    { 
      nombre: 'turno', 
      label: 'Turno', 
      tipo: 'select',
      opciones: ['DIA', 'NOCHE']
    },
    { 
      nombre: 'activo', 
      label: 'Activo', 
      tipo: 'select',
      opciones: ['SI', 'NO']
    }
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

  procesarExcelGeneral(event: any) {
  if (!this.modalContenido) {
    return;
  }

  switch (this.modalContenido.tipo) {
    case 'Equipo':
      this.procesarExcelEquipo(event);
      break;

    case 'Toneladas':
      this.procesarExcelToneladas(event);
      break;

    case 'OrigenDestino':      // 👈 necesario
      this.procesarExcelOrigenDestino(event);
      break;

    // 🔹 Agrega aquí otros tipos cuando implementes sus funciones
    default:
      break;
  }
}

    editarDato(dato: any, index: number) {
  this.editando = true;
  this.indiceEditando = index;
  this.datoOriginal = {...dato};
  
  // Clonamos el dato para editarlo
  this.nuevoDato = {...dato};
}
  
  actualizarDatos() {
  if (Object.values(this.nuevoDato).some(val => val !== '')) {
    const datosActualizados = {...this.nuevoDato};
    const id = this.modalContenido.datos[this.indiceEditando].id;

    if (this.modalContenido.tipo === 'Tipo de Perforación') {
      if (datosActualizados.permitido_medicion === 'SI') {
        datosActualizados.permitido_medicion = 1;
      } else if (datosActualizados.permitido_medicion === 'NO') {
        datosActualizados.permitido_medicion = 0;
      }

      this.tipoPerforacionService.updateTipoPerforacion(id, datosActualizados).subscribe({
        next: (data) => {
          data.permitido_medicion = data.permitido_medicion === 1 ? 'SI' : 'NO';
          this.modalContenido.datos[this.indiceEditando] = data;
          this.cancelarEdicion();
        },
        error: (err) => console.error('Error al actualizar:', err)
      });
    }
    else if (this.modalContenido.tipo === 'Equipo') {
      this.equipoService.updateEquipo(id, datosActualizados).subscribe({
        next: (data) => {
          this.modalContenido.datos[this.indiceEditando] = data;
          this.cancelarEdicion();
        },
        error: (err) => console.error('Error al actualizar:', err)
      });
    }else if (this.modalContenido.tipo === 'OrigenDestino') {
  this.origenDestinoService.updateOrigenDestino(id, datosActualizados).subscribe({
    next: (data) => {
      this.modalContenido.datos[this.indiceEditando] = data;
      this.cancelarEdicion();
    },
    error: (err) => console.error('Error al actualizar Origen/Destino:', err)
  });
}else if (this.modalContenido.tipo === 'Acero') {
      this.procesoAceroService.updateProceso(id, datosActualizados).subscribe({
        next: (data) => {
          this.modalContenido.datos[this.indiceEditando] = data;
          this.cancelarEdicion();
        }
      });
    }

  }
}

cancelarEdicion() {
  this.editando = false;
  this.indiceEditando = -1;
  this.nuevoDato = {};
  this.datoOriginal = null;
}

procesarExcelOrigenDestino(event: any) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e: any) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const excelData: any[] = XLSX.utils.sheet_to_json(sheet, { raw: false });

    const registros = excelData.map(row => ({
      operacion: row["OPERACION"] || null,
      tipo: row["TIPO"] || null,
      nombre: row["NOMBRE"] || null
    }));

    this.cerrarModal();
    const dialogRef = this.mostrarPantallaCarga();

    registros.forEach(nuevoRegistro => {
      this.origenDestinoService.createOrigenDestino(nuevoRegistro).subscribe({
        next: (data) => {
          this.modalContenido.datos.push(data);
        },
        error: (err) => console.error('Error al guardar Origen/Destino:', err)
      });
    });

    this.dialog.closeAll();
  };

  reader.readAsArrayBuffer(file);
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

    procesarExcelToneladas(event: any) {
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

    const excelData: any[] = XLSX.utils.sheet_to_json(sheet, { raw: false });

    const toneladas = excelData.map(row => ({
      fecha: row["FECHA"] || null,
      turno: row["TURNO"] || null,
      zona: row["ZONA"] || null,
      tipo: row["TIPO"] || null,
      labor: row["LABOR"] || null,
       toneladas: row["TONELADAS"] ? Number(row["TONELADAS"]) : 0
    }));

    this.cerrarModal();
    const dialogRef = this.mostrarPantallaCarga();

    // Procesar cada tonelada individualmente con el service
    toneladas.forEach(nuevoRegistro => {
      this.toneladasService.createTonelada(nuevoRegistro).subscribe({
        next: (data) => {
          this.modalContenido.datos.push(data);
        },
        error: (err) => console.error('Error al guardar Toneladas:', err)
      });
    });

    this.dialog.closeAll();
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

    this.nuevoDato = {};
  if (button.campos) {
    button.campos.forEach((campo: { valorPorDefecto: undefined; nombre: string | number; }) => {
      if (campo.valorPorDefecto !== undefined) {
        this.nuevoDato[campo.nombre] = campo.valorPorDefecto;
      } else {
        this.nuevoDato[campo.nombre] = ''; // inicializa vacío
      }
    });
  }
  
    if (button.tipo === 'Tipo de Perforación') {
  this.tipoPerforacionService.getTiposPerforacion().subscribe({
    next: (data) => {
      // Mapear 1 -> 'SI' y 0 -> 'NO' antes de asignar
      this.modalContenido.datos = data.map(item => ({
        ...item,
        permitido_medicion: item.permitido_medicion === 1 ? 'SI' : 'NO'
      }));
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
    }else if (button.tipo === 'Toneladas') {
  this.toneladasService.getToneladas().subscribe({
    next: (data) => {
      this.modalContenido.datos = data;
    },
    error: (err) => console.error('Error al cargar Toneladas:', err)
  });
}else if (button.tipo === 'OrigenDestino') {
  this.origenDestinoService.getOrigenesDestinos().subscribe({
    next: (data) => {
      this.modalContenido.datos = data;
    },
    error: (err) => console.error('Error al cargar Origen/Destino:', err)
  });
}else if (button.tipo === 'Acero') {
    this.procesoAceroService.getProcesos().subscribe({
      next: (data) => {
        this.modalContenido.datos = data;
      },
      error: (err) => console.error('Error al cargar Proceso Acero:', err)
    });
  }else if (button.tipo === 'OperadorAcero') {
    this.operadorAceroService.getOperadores().subscribe({
      next: (data) => {
        this.modalContenido.datos = data.map(o => ({
  ...o,
  activo: o.activo == 1 ? 'SI' : 'NO'   // <= con == permite comparar "1" y 1
}));

      }
    });
  }else if (button.tipo === 'JefeGuardiaAcero') {
    this.jefeGuardiaAceroService.getJefes().subscribe({
      next: (data) => {
        this.modalContenido.datos = data.map(j => ({
  ...j,
  activo: j.activo == 1 ? 'SI' : 'NO'   // <= con == permite comparar "1" y 1
}));

      }
    });
  }


  }

  guardarDatos() {
    if (Object.values(this.nuevoDato).some(val => val !== '')) {
      const nuevoRegistro = { ...this.nuevoDato };


      if (this.modalContenido.tipo === 'Tipo de Perforación') {
  if (nuevoRegistro.permitido_medicion === 'SI') {
    nuevoRegistro.permitido_medicion = 1;
  } else if (nuevoRegistro.permitido_medicion === 'NO') {
    nuevoRegistro.permitido_medicion = 0;
  }

  this.tipoPerforacionService.createTipoPerforacion(nuevoRegistro).subscribe({
    next: (data) => {
      // Mapear el campo antes de insertar en la tabla
      data.permitido_medicion = data.permitido_medicion === 1 ? 'SI' : 'NO';
      this.modalContenido.datos.push(data);
    },
    error: (err) => console.error('Error al guardar Tipo de Perforación:', err)
  });
}
 else if (this.modalContenido.tipo === 'Equipo') {
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
      else if (this.modalContenido.tipo === 'Toneladas') {
  this.toneladasService.createTonelada(nuevoRegistro).subscribe({
    next: (data) => {
      this.modalContenido.datos.push(data);
    },
    error: (err) => console.error('Error al guardar Toneladas:', err)
  });
}else if (this.modalContenido.tipo === 'OrigenDestino') {
  this.origenDestinoService.createOrigenDestino(nuevoRegistro).subscribe({
    next: (data) => {
      this.modalContenido.datos.push(data);
    },
    error: (err) => console.error('Error al guardar Origen/Destino:', err)
  });
}else if (this.modalContenido.tipo === 'Acero') {
      this.procesoAceroService.createProceso(nuevoRegistro).subscribe({
        next: (data) => {
          this.modalContenido.datos.push(data);
        }
      });
    }else if (this.modalContenido.tipo === 'JefeGuardiaAcero') {
  // convertir antes de mandar a la API
  nuevoRegistro.activo = nuevoRegistro.activo === 'SI' ? 1 : 0;

  this.jefeGuardiaAceroService.createJefe(nuevoRegistro).subscribe({
    next: (data) => {
      // hacemos una copia SOLO para mostrar
      const dataConTexto = {
        ...data,
        activo: data.activo === 1 ? 'SI' : 'NO'
      };
      this.modalContenido.datos.push(dataConTexto);
    }
  });
} else if (this.modalContenido.tipo === 'OperadorAcero') {
  nuevoRegistro.activo = nuevoRegistro.activo === 'SI' ? 1 : 0;

  this.operadorAceroService.createOperador(nuevoRegistro).subscribe({
    next: (data) => {
      const dataConTexto = {
        ...data,
        activo: data.activo === 1 ? 'SI' : 'NO'
      };
      this.modalContenido.datos.push(dataConTexto);
    }
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
    }else if (this.modalContenido.tipo === 'Toneladas') {
  this.toneladasService.deleteTonelada(item.id).subscribe({
    next: () => {
      this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
    },
    error: (err) => console.error('Error al eliminar Tonelada:', err)
  });
}else if (this.modalContenido.tipo === 'OrigenDestino') {
  this.origenDestinoService.deleteOrigenDestino(item.id).subscribe({
    next: () => {
      this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
    },
    error: (err) => console.error('Error al eliminar Origen/Destino:', err)
  });
}else if (this.modalContenido.tipo === 'Acero') {
    this.procesoAceroService.deleteProceso(item.id).subscribe({
      next: () => {
        this.modalContenido.datos = this.modalContenido.datos.filter((dato: any) => dato.id !== item.id);
      }
    });
  }else if (this.modalContenido.tipo === 'JefeGuardiaAcero') {
    this.jefeGuardiaAceroService.deleteJefe(item.id).subscribe({
      next: () => this.modalContenido.datos = this.modalContenido.datos.filter((d: any) => d.id !== item.id)
    });
  } else if (this.modalContenido.tipo === 'OperadorAcero') {
    this.operadorAceroService.deleteOperador(item.id).subscribe({
      next: () => this.modalContenido.datos = this.modalContenido.datos.filter((d: any) => d.id !== item.id)
    });
  }

  


  }

  descargar(item: any): void {}
}