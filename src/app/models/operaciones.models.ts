// Operación principal
export interface NubeOperacion {
  id: number;
  turno: string;
  equipo: string;
  codigo: string;
  empresa: string;
  fecha: string;
  tipo_operacion: string;
  estado: string;
  envio: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relaciones
  horometros?: NubeHorometros[];
  checklists?: NubeCheckListOperacion[];
  estados?: NubeEstado[];
}

// Checklist Operación
export interface NubeCheckListOperacion {
  id: number;
  operacion_id: number;
  descripcion: string;
  decision: number;
  observacion: string;
  categoria: string;
  createdAt?: string;
  updatedAt?: string;
  
  // Relación
  operacion?: NubeOperacion;
}

// Horómetros
export interface NubeHorometros {
  id: number;
  operacion_id: number;
  nombre: string;
  inicial: number;
  final: number;
  EstaOP: number;
  EstaINOP: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relación
  operacion?: NubeOperacion;
}

// Estado
export interface NubeEstado {
  id: number;
  operacion_id: number;
  numero: number;
  estado: string;
  codigo: string;
  hora_inicio: string;
  hora_final: string;
  createdAt?: string;
  updatedAt?: string;
  
  // Relaciones
  operacion?: NubeOperacion;
  perforaciones_taladro_largo?: NubePerforacionTaladroLargo[];
  perforaciones_horizontal?: NubePerforacionHorizontal[];
  sostenimientos?: NubeSostenimiento[];
  carguios?: NubeCarguio[]; 
}

// Perforación Taladro Largo
export interface NubePerforacionTaladroLargo {
  id: number;
  zona: string;
  tipo_labor: string;
  labor: string;
  veta: string;
  nivel: string;
  tipo_perforacion: string;
  observacion?: string;
  estado_id: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relaciones
  estado?: NubeEstado;
  inter_perforaciones?: NubeInterPerforacionTaladroLargo[];
}

// Inter Perforación Taladro Largo
export interface NubeInterPerforacionTaladroLargo {
  id: number;
  codigo_actividad: string;
  nivel: string;
  tajo: string;
  nbroca: number;
  ntaladro: number;
  material: string;
  nbarras: number;
  longitud_perforacion: number;
  angulo_perforacion: number;
  nfilas_de_hasta: string;
  detalles_trabajo_realizado: string;
  perforaciontaladrolargo_id: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relación
  perforacion_taladro_largo?: NubePerforacionTaladroLargo;
}

// Perforación Horizontal
export interface NubePerforacionHorizontal {
  id: number;
  zona: string;
  tipo_labor: string;
  labor: string;
  veta: string;
  nivel: string;
  tipo_perforacion: string;
  observacion?: string;
  estado_id: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relaciones
  estado?: NubeEstado;
  inter_perforaciones_horizontal?: NubeInterPerforacionHorizontal[];
}

// Inter Perforación Horizontal
export interface NubeInterPerforacionHorizontal {
  id: number;
  codigo_actividad: string;
  nivel: string;
  labor: string;
  seccion_la_labor: string;
  nbroca: number;
  ntaladro: number;
  material: string;
  ntaladros_rimados: number;
  longitud_perforacion: number;
   metros_perforados?: number;
  detalles_trabajo_realizado: string;
  perforacionhorizontal_id: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relación
  perforacion_horizontal?: NubePerforacionHorizontal;
}

// Sostenimiento
export interface NubeSostenimiento {
  id: number;
  zona: string;
  tipo_labor: string;
  labor: string;
  veta: string;
  nivel: string;
  tipo_perforacion: string;
  observacion?: string;
  estado_id: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relaciones
  estado?: NubeEstado;
  inter_sostenimientos?: NubeInterSostenimiento[];
}

// Inter Sostenimiento
export interface NubeInterSostenimiento {
  id: number;
  codigo_actividad: string;
  nivel: string;
  labor: string;
  seccion_de_labor: string;
  nbroca: number;
  ntaladro: number;
  material: string;
  longitud_perforacion: number;
  malla_instalada: string;
   metros_perforados?: number;
   detalles_trabajo_realizado: string;
  sostenimiento_id: number;
  createdAt?: string;
  updatedAt?: string;
  
  // Relación
  sostenimiento?: NubeSostenimiento;
}

// Carguio
export interface NubeCarguio {
  id: number;
  estado_id: number;
  tipo_labor?: string;          
  labor?: string;               
  tipo_labor_manual?: string;   
  labor_manual?: string;        
  ncucharas?: number;           
  observacion?: string;         
  createdAt?: string;
  updatedAt?: string;

  // Relación
  estado?: NubeEstado;
}
